//using Azure.Storage.Blobs;
//using Microsoft.Extensions.Azure;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Security.Principal;

namespace ManagedIdentityPlugin
{
    /// <summary>
    /// Plugin development guide: https://docs.microsoft.com/powerapps/developer/common-data-service/plug-ins
    /// Best practices and guidance: https://docs.microsoft.com/powerapps/developer/common-data-service/best-practices/business-logic/
    /// </summary>
    public class Flow : PluginBase
    {
        public Flow(string unsecureConfiguration, string secureConfiguration)
            : base(typeof(Flow))
        {
            // TODO: Implement your custom configuration handling
            // https://docs.microsoft.com/powerapps/developer/common-data-service/register-plug-in#set-configuration-data
        }

        // Entry point for custom business logic execution
        protected override void ExecuteDataversePlugin(ILocalPluginContext localPluginContext)
        {
            if (localPluginContext == null)
            {
                throw new ArgumentNullException(nameof(localPluginContext));
            }

            var context = localPluginContext.PluginExecutionContext;
            var identityService = (IManagedIdentityService)localPluginContext.ServiceProvider.GetService(typeof(IManagedIdentityService));
            var scopes = new List<string> { "https://service.flow.microsoft.com/.default" };
            var token = identityService.AcquireToken(scopes);
            localPluginContext.TracingService.Trace(token);
        }
    }
}
