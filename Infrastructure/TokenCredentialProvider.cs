using Azure.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ManagedIdentityPlugin
{
    internal class TokenCredentialProvider : TokenCredential
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
