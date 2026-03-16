using Microsoft.Xrm.Sdk;
using System;

namespace ManagedIdentityPlugin
{
    public class SharePointEmbeddedDeleteContainer : PluginBase
    {
        public SharePointEmbeddedDeleteContainer(string unsecureConfiguration, string secureConfiguration) : base(typeof(SharePointEmbeddedDeleteContainer))
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
                client.DeleteContainer(containerId);
            }

            context.OutputParameters["ContainerId"] = containerId;
            context.OutputParameters["Status"] = "Deleted";
        }
    }
}