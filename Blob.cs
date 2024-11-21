using Azure.Core;
using Azure.Storage.Blobs;
using Azure.Storage.Sas;
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
        public Blob(string unsecureConfiguration, string secureConfiguration) : base(typeof(Blob))
        {
        }

        protected override void ExecuteDataversePlugin(ILocalPluginContext localPluginContext)
        {
            if (localPluginContext == null)
            {
                throw new ArgumentNullException(nameof(localPluginContext));
            }

            var blobUrl = "https://<blobname>.blob.core.windows.net";

            var identityService = (IManagedIdentityService)localPluginContext.ServiceProvider.GetService(typeof(IManagedIdentityService));
            var scopes = new List<string> { "https://storage.azure.com/.default" };
            var token = identityService.AcquireToken(scopes);
            var blobTokenProvider = new TokenCredentialProvider(token);

            BlobServiceClient client = new BlobServiceClient(new Uri(blobUrl), blobTokenProvider);

            GenerateSaSToken(localPluginContext, client);

            IterateContainers(localPluginContext, client);
        }

        private static void IterateContainers(ILocalPluginContext localPluginContext, BlobServiceClient client)
        {
            var containers = client.GetBlobContainers();
            foreach (var container in containers)
            {
                localPluginContext.TracingService.Trace(container.Name);
            }
        }

        private static void GenerateSaSToken(ILocalPluginContext localPluginContext, BlobServiceClient client)
        {
            var userDelegationKey = client.GetUserDelegationKey(DateTimeOffset.UtcNow, DateTimeOffset.UtcNow.AddDays(1));

            var blobContainerClient = client.GetBlobContainerClient("plugin");  
            var blobClient = blobContainerClient.GetBlobClient("image.jpg");

            // Get a user delegation key 
            var sasBuilder = new BlobSasBuilder()
            {
                BlobContainerName = blobClient.BlobContainerName,
                BlobName = blobClient.Name,
                Resource = "b", // b for blob, c for container
                StartsOn = DateTimeOffset.UtcNow,
                ExpiresOn = DateTimeOffset.UtcNow.AddHours(4),
            };
            sasBuilder.SetPermissions(BlobSasPermissions.Read | BlobSasPermissions.Write);

            string sasToken = sasBuilder.ToSasQueryParameters(userDelegationKey, client.AccountName).ToString();

            localPluginContext.TracingService.Trace("SAS-Token {0}", sasToken);
            localPluginContext.TracingService.Trace($"{blobClient.Uri}?{sasToken}");
        }

        public class TokenCredentialProvider : TokenCredential
        {
            private string _token;

            public TokenCredentialProvider(string token)
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
