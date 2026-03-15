using Microsoft.Xrm.Sdk;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;

namespace ManagedIdentityPlugin
{
    internal static class PluginParameterHelper
    {
        public static string GetRequiredString(IPluginExecutionContext context, string parameterName)
        {
            if (!context.InputParameters.Contains(parameterName) || context.InputParameters[parameterName] == null)
            {
                throw new InvalidPluginExecutionException($"Input parameter '{parameterName}' is required.");
            }

            var value = context.InputParameters[parameterName] as string;
            if (string.IsNullOrWhiteSpace(value))
            {
                throw new InvalidPluginExecutionException($"Input parameter '{parameterName}' must be a non-empty string.");
            }

            return value;
        }

        public static string GetOptionalString(IPluginExecutionContext context, string parameterName)
        {
            if (!context.InputParameters.Contains(parameterName) || context.InputParameters[parameterName] == null)
            {
                return null;
            }

            return context.InputParameters[parameterName] as string;
        }

        public static string GetRequiredString(IPluginExecutionContext context, params string[] parameterNames)
        {
            foreach (var parameterName in parameterNames)
            {
                var value = GetOptionalString(context, parameterName);
                if (!string.IsNullOrWhiteSpace(value))
                {
                    return value;
                }
            }

            throw new InvalidPluginExecutionException($"One of the following input parameters is required: {string.Join(", ", parameterNames)}.");
        }
    }

    internal sealed class SharePointEmbeddedGraphClient : IDisposable
    {
        private const string GraphScope = "https://graph.microsoft.com/.default";
        private const string GraphBaseUrl = "https://graph.microsoft.com/v1.0";
        private static readonly HashSet<string> ContainerRoles = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "reader",
            "writer",
            "manager",
            "owner"
        };

        private readonly HttpClient _httpClient;
        private readonly ILocalPluginContext _localPluginContext;

        public SharePointEmbeddedGraphClient(ILocalPluginContext localPluginContext)
            : this(localPluginContext, null)
        {
        }

        public SharePointEmbeddedGraphClient(ILocalPluginContext localPluginContext, string accessToken)
        {
            if (localPluginContext == null)
            {
                throw new ArgumentNullException(nameof(localPluginContext));
            }

            _localPluginContext = localPluginContext;

            var token = accessToken;
            if (string.IsNullOrWhiteSpace(token))
            {
                var identityService = (IManagedIdentityService)localPluginContext.ServiceProvider.GetService(typeof(IManagedIdentityService));
                if (identityService == null)
                {
                    throw new InvalidPluginExecutionException("Managed identity service is not available in the plugin service provider.");
                }

                token = identityService.AcquireToken(new List<string> { GraphScope });
            }

            _httpClient = new HttpClient();
            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
            _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        }

        public CreateContainerResult CreateContainer(string containerTypeId, string displayName, string description)
        {
            var requestBody = new JObject
            {
                ["containerTypeId"] = containerTypeId,
                ["displayName"] = displayName
            };

            if (!string.IsNullOrWhiteSpace(description))
            {
                requestBody["description"] = description;
            }

            var createdContainer = SendJson(HttpMethod.Post, "/storage/fileStorage/containers", requestBody);
            var containerId = createdContainer.Value<string>("id");
            var drive = GetDrive(containerId);

            return new CreateContainerResult
            {
                ContainerId = containerId,
                ContainerTypeId = createdContainer.Value<string>("containerTypeId"),
                DriveId = drive.DriveId,
                DisplayName = createdContainer.Value<string>("displayName"),
                Description = createdContainer.Value<string>("description"),
                Status = createdContainer.Value<string>("status"),
                WebUrl = drive.WebUrl
            };
        }

        public ContainerProvisionResult CreateContainerWithFile(
            string containerTypeId,
            string displayName,
            string description,
            string fileName,
            string fileContentBase64,
            string contentType,
            IReadOnlyCollection<ContainerAccessAssignment> assignments)
        {
            var createdContainer = CreateContainer(containerTypeId, displayName, description);
            var grantedPermissions = new List<ContainerPermissionResult>();

            if (assignments != null)
            {
                foreach (var assignment in assignments)
                {
                    if (assignment == null || string.IsNullOrWhiteSpace(assignment.UserPrincipalName))
                    {
                        continue;
                    }

                    grantedPermissions.Add(GrantContainerPermission(createdContainer.ContainerId, assignment.UserPrincipalName, assignment.Role));
                }
            }

            var uploadedFile = UploadFile(createdContainer.ContainerId, null, fileName, fileContentBase64, contentType);

            return new ContainerProvisionResult
            {
                Container = createdContainer,
                File = uploadedFile,
                Permissions = grantedPermissions
            };
        }

        public UploadFileResult UploadFile(string containerId, string folderPath, string fileName, string fileContentBase64, string contentType)
        {
            byte[] fileBytes;
            try
            {
                fileBytes = Convert.FromBase64String(fileContentBase64);
            }
            catch (FormatException ex)
            {
                throw new InvalidPluginExecutionException("Input parameter 'FileContentBase64' must be valid base64.", ex);
            }

            var drive = GetDrive(containerId);
            var normalizedFolderPath = NormalizePath(folderPath);

            if (!string.IsNullOrWhiteSpace(normalizedFolderPath))
            {
                EnsureFolderPath(drive.DriveId, normalizedFolderPath);
            }

            var relativePath = CombinePath(normalizedFolderPath, fileName);
            var uploadUrl = string.Format("/drives/{0}/root:/{1}:/content", EscapeSegment(drive.DriveId), EncodePath(relativePath));

            using (var content = new ByteArrayContent(fileBytes))
            {
                content.Headers.ContentType = MediaTypeHeaderValue.Parse(string.IsNullOrWhiteSpace(contentType) ? "application/octet-stream" : contentType);

                var uploadedItem = SendJson(HttpMethod.Put, uploadUrl, content);
                return new UploadFileResult
                {
                    DriveId = drive.DriveId,
                    DriveItemId = uploadedItem.Value<string>("id"),
                    Name = uploadedItem.Value<string>("name"),
                    Size = uploadedItem.Value<long?>("size") ?? fileBytes.LongLength,
                    WebUrl = uploadedItem.Value<string>("webUrl")
                };
            }
        }

        public ContainerPermissionResult GrantContainerPermission(string containerId, string userPrincipalName, string role)
        {
            var normalizedRole = NormalizeContainerRole(role);
            var existingPermission = FindContainerPermission(containerId, userPrincipalName);
            if (existingPermission != null)
            {
                if (string.Equals(existingPermission.Role, normalizedRole, StringComparison.OrdinalIgnoreCase))
                {
                    return existingPermission;
                }

                DeleteContainerPermission(containerId, existingPermission.PermissionId);
            }

            var requestBody = new JObject
            {
                ["roles"] = new JArray(normalizedRole),
                ["grantedToV2"] = new JObject
                {
                    ["user"] = new JObject
                    {
                        ["userPrincipalName"] = userPrincipalName
                    }
                }
            };

            var createdPermission = SendJson(
                HttpMethod.Post,
                string.Format("/storage/fileStorage/containers/{0}/permissions", EscapeSegment(containerId)),
                requestBody);

            return ParseContainerPermission(createdPermission);
        }

        public ContainerPermissionResult RevokeContainerPermission(string containerId, string permissionId, string userPrincipalName)
        {
            ContainerPermissionResult existingPermission;

            if (!string.IsNullOrWhiteSpace(permissionId))
            {
                existingPermission = GetContainerPermission(containerId, permissionId);
            }
            else
            {
                existingPermission = FindContainerPermission(containerId, userPrincipalName);
            }

            if (existingPermission == null)
            {
                throw new InvalidPluginExecutionException("The specified container permission was not found.");
            }

            DeleteContainerPermission(containerId, existingPermission.PermissionId);
            return existingPermission;
        }

        public List<ContainerSummaryResult> ListContainers(string containerTypeId, int? top)
        {
            if (string.IsNullOrWhiteSpace(containerTypeId))
            {
                throw new InvalidPluginExecutionException("Input parameter 'ContainerTypeId' is required.");
            }

            Guid parsedContainerTypeId;
            if (!Guid.TryParse(containerTypeId, out parsedContainerTypeId))
            {
                throw new InvalidPluginExecutionException("Input parameter 'ContainerTypeId' must be a valid GUID.");
            }

            var relativeUrl = string.Format(
                "/storage/fileStorage/containers?$filter=containerTypeId eq {0}",
                parsedContainerTypeId);

            if (top.HasValue && top.Value > 0)
            {
                relativeUrl += string.Format("&$top={0}", top.Value);
            }

            var response = SendJson(HttpMethod.Get, relativeUrl, (JToken)null);
            var containers = response["value"] as JArray;
            var results = new List<ContainerSummaryResult>();
            var deletedContainerIds = ListDeletedContainerIds(containerTypeId);

            if (containers == null)
            {
                return results;
            }

            foreach (var container in containers.OfType<JObject>())
            {
                var containerId = container.Value<string>("id");
                if (deletedContainerIds.Contains(containerId))
                {
                    continue;
                }

                results.Add(BuildContainerSummary(container));
            }

            return results;
        }

        public List<ContainerSummaryResult> ListDeletedContainers(string containerTypeId, int? top)
        {
            if (string.IsNullOrWhiteSpace(containerTypeId))
            {
                throw new InvalidPluginExecutionException("Input parameter 'ContainerTypeId' is required.");
            }

            Guid parsedContainerTypeId;
            if (!Guid.TryParse(containerTypeId, out parsedContainerTypeId))
            {
                throw new InvalidPluginExecutionException("Input parameter 'ContainerTypeId' must be a valid GUID.");
            }

            var relativeUrl = string.Format(
                "/storage/fileStorage/deletedContainers?$filter=containerTypeId eq {0}",
                parsedContainerTypeId);

            if (top.HasValue && top.Value > 0)
            {
                relativeUrl += string.Format("&$top={0}", top.Value);
            }

            var response = SendJson(HttpMethod.Get, relativeUrl, (JToken)null);
            var containers = response["value"] as JArray;
            var results = new List<ContainerSummaryResult>();

            if (containers == null)
            {
                return results;
            }

            foreach (var container in containers.OfType<JObject>())
            {
                results.Add(BuildDeletedContainerSummary(container));
            }

            return results;
        }

        public ContainerSummaryResult GetContainerDetails(string containerId)
        {
            if (string.IsNullOrWhiteSpace(containerId))
            {
                throw new InvalidPluginExecutionException("Input parameter 'ContainerId' is required.");
            }

            var container = SendJson(
                HttpMethod.Get,
                string.Format("/storage/fileStorage/containers/{0}", EscapeSegment(containerId)),
                (JToken)null);

            return BuildContainerSummary(container);
        }

        public void DeleteContainer(string containerId)
        {
            if (string.IsNullOrWhiteSpace(containerId))
            {
                throw new InvalidPluginExecutionException("Input parameter 'ContainerId' is required.");
            }

            using (var response = Send(
                string.Format("/storage/fileStorage/containers/{0}", EscapeSegment(containerId)),
                HttpMethod.Delete,
                null,
                false))
            {
            }
        }

        public ContainerSummaryResult RestoreContainer(string containerId)
        {
            if (string.IsNullOrWhiteSpace(containerId))
            {
                throw new InvalidPluginExecutionException("Input parameter 'ContainerId' is required.");
            }

            var restored = SendJson(
                HttpMethod.Post,
                string.Format("/storage/fileStorage/deletedContainers/{0}/restore", EscapeSegment(containerId)),
                (JToken)null);

            return BuildContainerSummary(restored);
        }

        private DriveInfo GetDrive(string containerId)
        {
            var drive = SendJson(HttpMethod.Get, string.Format("/storage/fileStorage/containers/{0}/drive", EscapeSegment(containerId)), (JToken)null);
            return new DriveInfo
            {
                DriveId = drive.Value<string>("id"),
                WebUrl = drive.Value<string>("webUrl")
            };
        }

        private void EnsureFolderPath(string driveId, string folderPath)
        {
            var segments = folderPath.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            var currentParentId = "root";
            var currentSegments = new List<string>();

            foreach (var segment in segments)
            {
                currentSegments.Add(segment);

                var existingFolder = TryGetJson(string.Format("/drives/{0}/root:/{1}", EscapeSegment(driveId), EncodePath(currentSegments)));
                if (existingFolder != null)
                {
                    currentParentId = existingFolder.Value<string>("id");
                    continue;
                }

                var createFolderBody = new JObject
                {
                    ["name"] = segment,
                    ["folder"] = new JObject(),
                    ["@microsoft.graph.conflictBehavior"] = "replace"
                };

                var createdFolder = SendJson(
                    HttpMethod.Post,
                    string.Format("/drives/{0}/items/{1}/children", EscapeSegment(driveId), EscapeSegment(currentParentId)),
                    createFolderBody);

                currentParentId = createdFolder.Value<string>("id");
            }
        }

        private ContainerSummaryResult BuildContainerSummary(JObject container)
        {
            var containerId = container.Value<string>("id");
            var drive = GetDrive(containerId);

            return new ContainerSummaryResult
            {
                ContainerId = containerId,
                ContainerTypeId = container.Value<string>("containerTypeId"),
                DriveId = drive.DriveId,
                DisplayName = container.Value<string>("displayName"),
                Description = container.Value<string>("description"),
                Status = container.Value<string>("status"),
                DeletedDateTime = container.Value<string>("deletedDateTime"),
                WebUrl = drive.WebUrl,
                Files = ListDriveItems(drive.DriveId),
                Permissions = ListContainerPermissions(containerId)
            };
        }

        private static ContainerSummaryResult BuildDeletedContainerSummary(JObject container)
        {
            return new ContainerSummaryResult
            {
                ContainerId = container.Value<string>("id"),
                ContainerTypeId = container.Value<string>("containerTypeId"),
                DisplayName = container.Value<string>("displayName"),
                Description = container.Value<string>("description"),
                Status = container.Value<string>("status"),
                DeletedDateTime = container.Value<string>("deletedDateTime"),
                Files = new List<ContainerFileResult>(),
                Permissions = new List<ContainerPermissionResult>()
            };
        }

        private HashSet<string> ListDeletedContainerIds(string containerTypeId)
        {
            Guid parsedContainerTypeId;
            if (!Guid.TryParse(containerTypeId, out parsedContainerTypeId))
            {
                return new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            }

            var response = SendJson(
                HttpMethod.Get,
                string.Format("/storage/fileStorage/deletedContainers?$filter=containerTypeId eq {0}", parsedContainerTypeId),
                (JToken)null);

            var values = response["value"] as JArray;
            if (values == null)
            {
                return new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            }

            return new HashSet<string>(
                values.OfType<JObject>()
                    .Select((container) => container.Value<string>("id"))
                    .Where((containerId) => !string.IsNullOrWhiteSpace(containerId)),
                StringComparer.OrdinalIgnoreCase);
        }

        private List<ContainerFileResult> ListDriveItems(string driveId)
        {
            var response = SendJson(
                HttpMethod.Get,
                string.Format(
                    "/drives/{0}/root/children?$select=id,name,size,webUrl,createdDateTime,lastModifiedDateTime,file,folder",
                    EscapeSegment(driveId)),
                (JToken)null);

            var items = new List<ContainerFileResult>();
            var values = response["value"] as JArray;
            if (values == null)
            {
                return items;
            }

            foreach (var item in values.OfType<JObject>())
            {
                items.Add(new ContainerFileResult
                {
                    DriveItemId = item.Value<string>("id"),
                    Name = item.Value<string>("name"),
                    Size = item.Value<long?>("size") ?? 0,
                    WebUrl = item.Value<string>("webUrl"),
                    CreatedDateTime = item.Value<string>("createdDateTime"),
                    LastModifiedDateTime = item.Value<string>("lastModifiedDateTime"),
                    ContentType = item.SelectToken("file.mimeType")?.Value<string>(),
                    IsFolder = item["folder"] != null
                });
            }

            return items;
        }

        private List<ContainerPermissionResult> ListContainerPermissions(string containerId)
        {
            var response = SendJson(
                HttpMethod.Get,
                string.Format("/storage/fileStorage/containers/{0}/permissions", EscapeSegment(containerId)),
                (JToken)null);

            var permissions = new List<ContainerPermissionResult>();
            var values = response["value"] as JArray;
            if (values == null)
            {
                return permissions;
            }

            foreach (var permission in values.OfType<JObject>())
            {
                permissions.Add(ParseContainerPermission(permission));
            }

            return permissions;
        }

        private ContainerPermissionResult FindContainerPermission(string containerId, string userPrincipalName)
        {
            if (string.IsNullOrWhiteSpace(userPrincipalName))
            {
                throw new InvalidPluginExecutionException("Input parameter 'UserPrincipalName' is required.");
            }

            var response = SendJson(
                HttpMethod.Get,
                string.Format("/storage/fileStorage/containers/{0}/permissions", EscapeSegment(containerId)),
                (JToken)null);

            var permissions = response["value"] as JArray;
            if (permissions == null)
            {
                return null;
            }

            foreach (var permission in permissions.OfType<JObject>())
            {
                var permissionUpn = permission.SelectToken("grantedToV2.user.userPrincipalName")?.Value<string>();
                if (string.Equals(permissionUpn, userPrincipalName, StringComparison.OrdinalIgnoreCase))
                {
                    return ParseContainerPermission(permission);
                }
            }

            return null;
        }

        private ContainerPermissionResult GetContainerPermission(string containerId, string permissionId)
        {
            var permission = SendJson(
                HttpMethod.Get,
                string.Format(
                    "/storage/fileStorage/containers/{0}/permissions/{1}",
                    EscapeSegment(containerId),
                    EscapeSegment(permissionId)),
                (JToken)null);

            return ParseContainerPermission(permission);
        }

        private void DeleteContainerPermission(string containerId, string permissionId)
        {
            using (var response = Send(
                string.Format(
                    "/storage/fileStorage/containers/{0}/permissions/{1}",
                    EscapeSegment(containerId),
                    EscapeSegment(permissionId)),
                HttpMethod.Delete,
                null,
                false))
            {
            }
        }

        private static ContainerPermissionResult ParseContainerPermission(JObject permission)
        {
            var roles = permission["roles"] as JArray;
            return new ContainerPermissionResult
            {
                PermissionId = permission.Value<string>("id"),
                Role = roles == null ? null : roles.Values<string>().FirstOrDefault(),
                UserPrincipalName = permission.SelectToken("grantedToV2.user.userPrincipalName")?.Value<string>()
            };
        }

        private static string NormalizeContainerRole(string role)
        {
            if (string.IsNullOrWhiteSpace(role))
            {
                throw new InvalidPluginExecutionException("Input parameter 'Role' is required.");
            }

            var normalizedRole = role.Trim().ToLowerInvariant();
            if (!ContainerRoles.Contains(normalizedRole))
            {
                throw new InvalidPluginExecutionException("Input parameter 'Role' must be one of: reader, writer, manager, owner.");
            }

            return normalizedRole;
        }

        private JObject TryGetJson(string relativeUrl)
        {
            using (var response = Send(relativeUrl, HttpMethod.Get, null, true))
            {
                if (response.StatusCode == HttpStatusCode.NotFound)
                {
                    return null;
                }

                return ParseJson(response);
            }
        }

        private JObject SendJson(HttpMethod method, string relativeUrl, JToken body, string graphBaseUrl = GraphBaseUrl)
        {
            var content = body == null
                ? null
                : new StringContent(body.ToString(Formatting.None), Encoding.UTF8, "application/json");

            using (content)
            using (var response = Send(relativeUrl, method, content, false, graphBaseUrl))
            {
                return ParseJson(response);
            }
        }

        private JObject SendJson(HttpMethod method, string relativeUrl, HttpContent content, string graphBaseUrl = GraphBaseUrl)
        {
            using (var response = Send(relativeUrl, method, content, false, graphBaseUrl))
            {
                return ParseJson(response);
            }
        }

        private HttpResponseMessage Send(string relativeUrl, HttpMethod method, HttpContent content, bool allowNotFound, string graphBaseUrl = GraphBaseUrl)
        {
            var request = new HttpRequestMessage(method, graphBaseUrl + relativeUrl);
            if (content != null)
            {
                request.Content = content;
            }

            _localPluginContext.Trace($"Graph request {method} {graphBaseUrl}{relativeUrl}");

            var response = _httpClient.SendAsync(request).GetAwaiter().GetResult();
            if (allowNotFound && response.StatusCode == HttpStatusCode.NotFound)
            {
                return response;
            }

            if (response.IsSuccessStatusCode)
            {
                return response;
            }

            var body = response.Content == null
                ? string.Empty
                : response.Content.ReadAsStringAsync().GetAwaiter().GetResult();

            throw new InvalidPluginExecutionException(
                string.Format(
                    "Microsoft Graph request failed. Status: {0} {1}. Body: {2}",
                    (int)response.StatusCode,
                    response.ReasonPhrase,
                    body));
        }

        private static JObject ParseJson(HttpResponseMessage response)
        {
            if (response.Content == null)
            {
                return new JObject();
            }

            var body = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
            if (string.IsNullOrWhiteSpace(body))
            {
                return new JObject();
            }

            return JObject.Parse(body);
        }

        private static string NormalizePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return null;
            }

            return path.Replace('\\', '/').Trim('/');
        }

        private static string CombinePath(string folderPath, string fileName)
        {
            return string.IsNullOrWhiteSpace(folderPath)
                ? fileName
                : string.Format("{0}/{1}", folderPath, fileName);
        }

        private static string EncodePath(IEnumerable<string> segments)
        {
            return string.Join("/", segments.Where(segment => !string.IsNullOrWhiteSpace(segment)).Select(Uri.EscapeDataString));
        }

        private static string EncodePath(string path)
        {
            return EncodePath(path.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries));
        }

        private static string EscapeSegment(string value)
        {
            return Uri.EscapeDataString(value);
        }

        private static string EscapeODataString(string value)
        {
            return (value ?? string.Empty).Replace("'", "''");
        }

        public void Dispose()
        {
            _httpClient.Dispose();
        }
    }

    internal sealed class CreateContainerResult
    {
        public string ContainerId { get; set; }

        public string ContainerTypeId { get; set; }

        public string DriveId { get; set; }

        public string DisplayName { get; set; }

        public string Description { get; set; }

        public string Status { get; set; }

        public string WebUrl { get; set; }
    }

    internal sealed class UploadFileResult
    {
        public string DriveId { get; set; }

        public string DriveItemId { get; set; }

        public string Name { get; set; }

        public long Size { get; set; }

        public string WebUrl { get; set; }
    }

    internal sealed class DriveInfo
    {
        public string DriveId { get; set; }

        public string WebUrl { get; set; }
    }

    internal sealed class ContainerPermissionResult
    {
        public string PermissionId { get; set; }

        public string UserPrincipalName { get; set; }

        public string Role { get; set; }
    }

    internal sealed class ContainerAccessAssignment
    {
        public string UserPrincipalName { get; set; }

        public string Role { get; set; }
    }

    internal sealed class ContainerFileResult
    {
        public string DriveItemId { get; set; }

        public string Name { get; set; }

        public long Size { get; set; }

        public string WebUrl { get; set; }

        public string CreatedDateTime { get; set; }

        public string LastModifiedDateTime { get; set; }

        public string ContentType { get; set; }

        public bool IsFolder { get; set; }
    }

    internal sealed class ContainerSummaryResult
    {
        public string ContainerId { get; set; }

        public string ContainerTypeId { get; set; }

        public string DriveId { get; set; }

        public string DisplayName { get; set; }

        public string Description { get; set; }

        public string Status { get; set; }

        public string DeletedDateTime { get; set; }

        public string WebUrl { get; set; }

        public List<ContainerFileResult> Files { get; set; }

        public List<ContainerPermissionResult> Permissions { get; set; }
    }

    internal sealed class ContainerProvisionResult
    {
        public CreateContainerResult Container { get; set; }

        public UploadFileResult File { get; set; }

        public List<ContainerPermissionResult> Permissions { get; set; }
    }
}