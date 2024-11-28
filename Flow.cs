using Azure.Core;
using Azure.Storage.Blobs;
using Azure.Storage.Sas;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;

namespace ManagedIdentityPlugin
{
    /// <summary>
    /// Plugin development guide: https://docs.microsoft.com/powerapps/developer/common-data-service/plug-ins
    /// Best practices and guidance: https://docs.microsoft.com/powerapps/developer/common-data-service/best-practices/business-logic/
    /// </summary>
    public class Flow : PluginBase
    {
        public Flow(string unsecureConfiguration, string secureConfiguration) : base(typeof(Flow))
        {
        }

        protected override void ExecuteDataversePlugin(ILocalPluginContext localPluginContext)
        {
            if (localPluginContext == null)
            {
                throw new ArgumentNullException(nameof(localPluginContext));
            }

            var flowUrl = "";

            var tracingService = localPluginContext.TracingService;
            var identityService = (IManagedIdentityService)localPluginContext.ServiceProvider.GetService(typeof(IManagedIdentityService));
            var scopes = new List<string> { "https://service.flow.microsoft.com//.default" };
            var token = identityService.AcquireToken(scopes);
            var client = new HttpClient();

            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

            var resp = client.GetAsync(flowUrl);

            tracingService.Trace(resp.Result.Content.ReadAsStringAsync().Result);
        }
    }
}
