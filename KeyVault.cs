using Azure.Core;
using Azure.Security.KeyVault.Secrets;
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
    public class KeyVault : PluginBase
    {
        public KeyVault(string unsecureConfiguration, string secureConfiguration) : base(typeof(KeyVault))
        {
        }

        protected override void ExecuteDataversePlugin(ILocalPluginContext localPluginContext)
        {
            if (localPluginContext == null)
            {
                throw new ArgumentNullException(nameof(localPluginContext));
            }

            var tracingService = localPluginContext.TracingService;

            var keyVaultName = "";

            var identityService = (IManagedIdentityService)localPluginContext.ServiceProvider.GetService(typeof(IManagedIdentityService));
            var scopes = new List<string> { "https://vault.azure.net/.default" };
            var token = identityService.AcquireToken(scopes);

            var kvUri = "https://" + keyVaultName + ".vault.azure.net";

            var client = new SecretClient(new Uri(kvUri), new TokenCredentialProvider(token));
            var secret = client.GetSecret("supersecret").Value.Value;
            tracingService.Trace($"Secret: {secret}");
        }
    }
}
