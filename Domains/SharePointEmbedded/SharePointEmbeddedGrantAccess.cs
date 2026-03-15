using Microsoft.Xrm.Sdk;
using System;

namespace ManagedIdentityPlugin
{
    public class SharePointEmbeddedGrantAccess : PluginBase
    {
        public SharePointEmbeddedGrantAccess(string unsecureConfiguration, string secureConfiguration) : base(typeof(SharePointEmbeddedGrantAccess))
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
            var userPrincipalName = PluginParameterHelper.GetRequiredString(context, "UserPrincipalName");
            var role = PluginParameterHelper.GetRequiredString(context, "Role");

            using (var client = new SharePointEmbeddedGraphClient(localPluginContext))
            {
                var result = client.GrantContainerPermission(containerId, userPrincipalName, role);

                context.OutputParameters["PermissionId"] = result.PermissionId;
                context.OutputParameters["UserPrincipalName"] = result.UserPrincipalName;
                context.OutputParameters["Role"] = result.Role;
            }
        }
    }
}