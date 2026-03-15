using Microsoft.Xrm.Sdk;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ManagedIdentityPlugin
{
    public class SharePointEmbeddedCreateContainerWithFile : PluginBase
    {
        public SharePointEmbeddedCreateContainerWithFile(string unsecureConfiguration, string secureConfiguration) : base(typeof(SharePointEmbeddedCreateContainerWithFile))
        {
        }

        protected override void ExecuteDataversePlugin(ILocalPluginContext localPluginContext)
        {
            if (localPluginContext == null)
            {
                throw new ArgumentNullException(nameof(localPluginContext));
            }

            var context = localPluginContext.PluginExecutionContext;
            var containerTypeId = PluginParameterHelper.GetRequiredString(context, "ContainerTypeId");
            var fileName = PluginParameterHelper.GetRequiredString(context, "FileName");
            var fileContentBase64 = PluginParameterHelper.GetRequiredString(context, "FileContentBase64");
            var contentType = PluginParameterHelper.GetOptionalString(context, "ContentType");
            var description = PluginParameterHelper.GetOptionalString(context, "Description");
            var permissionsJson = PluginParameterHelper.GetOptionalString(context, "PermissionsJson");
            var displayName = PluginParameterHelper.GetOptionalString(context, "DisplayName");

            if (string.IsNullOrWhiteSpace(displayName))
            {
                displayName = Path.GetFileNameWithoutExtension(fileName);
            }

            var assignments = ParseAssignments(permissionsJson);

            using (var client = new SharePointEmbeddedGraphClient(localPluginContext))
            {
                var provisioned = client.CreateContainerWithFile(
                    containerTypeId,
                    displayName,
                    description,
                    fileName,
                    fileContentBase64,
                    contentType,
                    assignments);

                var details = client.GetContainerDetails(provisioned.Container.ContainerId);
                var serializedPermissions = JsonConvert.SerializeObject(details.Permissions);
                var serializedContainer = JsonConvert.SerializeObject(details);

                context.OutputParameters["ContainerId"] = provisioned.Container.ContainerId;
                context.OutputParameters["DriveId"] = provisioned.Container.DriveId;
                context.OutputParameters["DriveItemId"] = provisioned.File.DriveItemId;
                context.OutputParameters["DisplayName"] = provisioned.Container.DisplayName;
                context.OutputParameters["Description"] = provisioned.Container.Description;
                context.OutputParameters["Status"] = provisioned.Container.Status;
                context.OutputParameters["WebUrl"] = provisioned.File.WebUrl ?? provisioned.Container.WebUrl;
                context.OutputParameters["FileName"] = provisioned.File.Name;
                context.OutputParameters["PermissionsJson"] = serializedPermissions;
                context.OutputParameters["ContainerJson"] = serializedContainer;
            }
        }

        private static IReadOnlyCollection<ContainerAccessAssignment> ParseAssignments(string permissionsJson)
        {
            if (string.IsNullOrWhiteSpace(permissionsJson))
            {
                return Array.Empty<ContainerAccessAssignment>();
            }

            JArray values;
            try
            {
                values = JArray.Parse(permissionsJson);
            }
            catch (JsonException ex)
            {
                throw new InvalidPluginExecutionException("Input parameter 'PermissionsJson' must be a valid JSON array.", ex);
            }

            var assignments = new Dictionary<string, ContainerAccessAssignment>(StringComparer.OrdinalIgnoreCase);
            foreach (var token in values.OfType<JObject>())
            {
                var userPrincipalName = token.Value<string>("UserPrincipalName") ?? token.Value<string>("userPrincipalName");
                var role = token.Value<string>("Role") ?? token.Value<string>("role");

                if (string.IsNullOrWhiteSpace(userPrincipalName) || string.IsNullOrWhiteSpace(role))
                {
                    throw new InvalidPluginExecutionException("Each entry in 'PermissionsJson' must contain userPrincipalName and role values.");
                }

                assignments[userPrincipalName] = new ContainerAccessAssignment
                {
                    UserPrincipalName = userPrincipalName,
                    Role = role
                };
            }

            return assignments.Values.ToList();
        }
    }
}