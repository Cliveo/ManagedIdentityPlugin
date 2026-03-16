using Microsoft.Xrm.Sdk;
using Newtonsoft.Json;
using System;

namespace ManagedIdentityPlugin
{
    public class SharePointEmbeddedRestoreContainer : PluginBase
    {
        public SharePointEmbeddedRestoreContainer(string unsecureConfiguration, string secureConfiguration) : base(typeof(SharePointEmbeddedRestoreContainer))
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

            using (var client = new SharePointEmbeddedGraphClient(localPluginContext))
            {
                var restored = client.RestoreContainer(containerId);
                context.OutputParameters["ContainerId"] = restored.ContainerId;
                context.OutputParameters["Status"] = "Restored";
                context.OutputParameters["ContainerJson"] = JsonConvert.SerializeObject(restored);
            }
        }
    }
}