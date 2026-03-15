using Microsoft.Xrm.Sdk;
using Newtonsoft.Json;
using System;

namespace ManagedIdentityPlugin
{
    public class SharePointEmbeddedGetContainerDetails : PluginBase
    {
        public SharePointEmbeddedGetContainerDetails(string unsecureConfiguration, string secureConfiguration) : base(typeof(SharePointEmbeddedGetContainerDetails))
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

            if (string.Equals(context.MessageName, "co_SharePointEmbeddedDeleteContainer", StringComparison.OrdinalIgnoreCase)
                || string.Equals(context.MessageName, "SharePointEmbeddedDeleteContainer", StringComparison.OrdinalIgnoreCase))
            {
                using (var client = new SharePointEmbeddedGraphClient(localPluginContext))
                {
                    client.DeleteContainer(containerId);
                }

                context.OutputParameters["ContainerId"] = containerId;
                context.OutputParameters["Status"] = "Deleted";
                return;
            }

            if (string.Equals(context.MessageName, "co_SharePointEmbeddedRestoreContainer", StringComparison.OrdinalIgnoreCase)
                || string.Equals(context.MessageName, "SharePointEmbeddedRestoreContainer", StringComparison.OrdinalIgnoreCase))
            {
                using (var client = new SharePointEmbeddedGraphClient(localPluginContext))
                {
                    var restored = client.RestoreContainer(containerId);
                    context.OutputParameters["ContainerId"] = restored.ContainerId;
                    context.OutputParameters["Status"] = "Restored";
                    context.OutputParameters["ContainerJson"] = JsonConvert.SerializeObject(restored);
                }

                return;
            }

            using (var client = new SharePointEmbeddedGraphClient(localPluginContext))
            {
                var details = client.GetContainerDetails(containerId);
                context.OutputParameters["ContainerJson"] = JsonConvert.SerializeObject(details);
            }
        }
    }
}