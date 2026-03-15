using Microsoft.Xrm.Sdk;
using Newtonsoft.Json;
using System;

namespace ManagedIdentityPlugin
{
    public class SharePointEmbeddedListDeletedContainers : PluginBase
    {
        public SharePointEmbeddedListDeletedContainers(string unsecureConfiguration, string secureConfiguration) : base(typeof(SharePointEmbeddedListDeletedContainers))
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
            var topValue = PluginParameterHelper.GetOptionalString(context, "Top");
            int parsedTop;
            int? top = int.TryParse(topValue, out parsedTop) && parsedTop > 0 ? parsedTop : (int?)null;

            using (var client = new SharePointEmbeddedGraphClient(localPluginContext))
            {
                var containers = client.ListDeletedContainers(containerTypeId, top);
                context.OutputParameters["ContainersJson"] = JsonConvert.SerializeObject(containers);
            }
        }
    }
}