using Azure.Core;
using Azure.Storage.Blobs;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Threading;
using System.Threading.Tasks;

namespace ManagedIdentityPlugin
{
    /// <summary>
    /// Plugin development guide: https://docs.microsoft.com/powerapps/developer/common-data-service/plug-ins
    /// Best practices and guidance: https://docs.microsoft.com/powerapps/developer/common-data-service/best-practices/business-logic/
    /// </summary>
    public class Blob : PluginBase
    {
        public Blob(string unsecureConfiguration, string secureConfiguration)
            : base(typeof(Blob))
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
            var scopes = new List<string> { "https://storage.azure.com/.default" };
            var token = identityService.AcquireToken(scopes);
            var blobTokenProvider = new BlobTokenProvider(token);
            localPluginContext.TracingService.Trace(token);
            var blobUrl = "https://ppolivergsa.blob.core.windows.net";

            BlobServiceClient client = new BlobServiceClient(new Uri(blobUrl), blobTokenProvider);
            var containers = client.GetBlobContainers();

            localPluginContext.TracingService.Trace($"Hello"); 
            localPluginContext.TracingService.Trace($"Accountz: {client.AccountName}"); 
            foreach (var container in containers) 
            {
                localPluginContext.TracingService.Trace(container.Name);
            }
            // TODO: Implement your custom business logic

            // Check for the entity on which the plugin would be registered
            //if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            //{
            //    var entity = (Entity)context.InputParameters["Target"];

            //    // Check for entity name on which this plugin would be registered
            //    if (entity.LogicalName == "account")
            //    {

            //    }
            //}
        }

        public class BlobTokenProvider : TokenCredential
        {
            private string _token;

            public BlobTokenProvider(string token)
            {
                _token = token;
            }
            public override ValueTask<AccessToken> GetTokenAsync(TokenRequestContext requestContext, CancellationToken cancellationToken)
            {
                return new ValueTask<AccessToken>(new AccessToken(_token, new DateTimeOffset(DateTime.UtcNow.AddMinutes(2))));
            }
            public override AccessToken GetToken(TokenRequestContext requestContext, CancellationToken cancellationToken)
            {
                return new AccessToken(_token, new DateTimeOffset(DateTime.UtcNow.AddMinutes(2)));
            }
        }
    }
}
