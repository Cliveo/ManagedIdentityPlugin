using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;

namespace ManagedIdentityPlugin
{
    public class GetFlowToken : PluginBase
    {
        public GetFlowToken(string unsecureConfiguration, string secureConfiguration) : base(typeof(GetFlowToken))
        {
        }

        protected override void ExecuteDataversePlugin(ILocalPluginContext localPluginContext)
        {
            if (localPluginContext == null)
            {
                throw new ArgumentNullException(nameof(localPluginContext));
            }

            var identityService = (IManagedIdentityService)localPluginContext.ServiceProvider.GetService(typeof(IManagedIdentityService));
            var scopes = new List<string> { "https://service.flow.microsoft.com//.default" };
            var token = identityService.AcquireToken(scopes);

            localPluginContext.PluginExecutionContext.OutputParameters["token"] = token;
        }
    }
}