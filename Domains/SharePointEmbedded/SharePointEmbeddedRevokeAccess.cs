using Microsoft.Xrm.Sdk;
using System;

namespace ManagedIdentityPlugin
{
    public class SharePointEmbeddedRevokeAccess : PluginBase
    {
        public SharePointEmbeddedRevokeAccess(string unsecureConfiguration, string secureConfiguration) : base(typeof(SharePointEmbeddedRevokeAccess))
        {
        }

        protected override void ExecuteDataversePlugin(ILocalPluginContext localPluginContext)
        {
            if (localPluginContext == null)
            {
                throw new ArgumentNullException(nameof(localPluginContext));
            }

            var context = localPluginContext.PluginExecutionContext;
            var containerId = PluginParameterHelper.GetRequiredString(context, "ContainerId");
            var permissionId = PluginParameterHelper.GetOptionalString(context, "PermissionId");
            var userPrincipalName = PluginParameterHelper.GetOptionalString(context, "UserPrincipalName");

            if (string.IsNullOrWhiteSpace(permissionId) && string.IsNullOrWhiteSpace(userPrincipalName))
            {
                throw new InvalidPluginExecutionException("Either 'PermissionId' or 'UserPrincipalName' must be supplied.");
            }

            using (var client = new SharePointEmbeddedGraphClient(localPluginContext))
            {
                var result = client.RevokeContainerPermission(containerId, permissionId, userPrincipalName);

                context.OutputParameters["PermissionId"] = result.PermissionId;
                context.OutputParameters["UserPrincipalName"] = result.UserPrincipalName;
                context.OutputParameters["Role"] = result.Role;
            }
        }
    }
}