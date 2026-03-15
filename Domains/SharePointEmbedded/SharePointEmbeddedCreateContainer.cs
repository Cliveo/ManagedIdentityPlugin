using Microsoft.Xrm.Sdk;
using System;

namespace ManagedIdentityPlugin
{
    public class SharePointEmbeddedCreateContainer : PluginBase
    {
        public SharePointEmbeddedCreateContainer(string unsecureConfiguration, string secureConfiguration) : base(typeof(SharePointEmbeddedCreateContainer))
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
            var displayName = PluginParameterHelper.GetRequiredString(context, "DisplayName");
            var description = PluginParameterHelper.GetOptionalString(context, "Description");

            using (var client = new SharePointEmbeddedGraphClient(localPluginContext))
            {
                var result = client.CreateContainer(containerTypeId, displayName, description);

                context.OutputParameters["ContainerId"] = result.ContainerId;
                context.OutputParameters["DriveId"] = result.DriveId;
                context.OutputParameters["DisplayName"] = result.DisplayName;
                context.OutputParameters["Description"] = result.Description;
                context.OutputParameters["Status"] = result.Status;
                context.OutputParameters["WebUrl"] = result.WebUrl;
            }
        }
    }
}