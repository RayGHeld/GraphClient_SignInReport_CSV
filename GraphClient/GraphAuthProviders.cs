using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GraphClient
{
    /// <summary>
    /// Refrence: https://github.com/microsoftgraph/msgraph-sdk-dotnet-auth
    /// </summary>
    class GraphAuthProviders
    {
        public class InteractiveAuthenticationProvider : IAuthenticationProvider
        {
            private static IPublicClientApplication Pca = null;
            private static List<String> Scopes = null;

            private InteractiveAuthenticationProvider()  
            {
                // intentionally left blank and private to prevent empty constructor
            }

            public InteractiveAuthenticationProvider( IPublicClientApplication pca, List<String> scopes )
            {
                Pca = pca;
                Scopes = scopes;
            }

            public async Task AuthenticateRequestAsync( HttpRequestMessage request )
            {
                IEnumerable<IAccount> accounts = null;
                AuthenticationResult result = null;

                accounts = await Pca.GetAccountsAsync();
                bool interactionRequired = false;

                try
                {
                    result = await Pca.AcquireTokenSilent( Scopes, accounts.FirstOrDefault() ).ExecuteAsync();
                }
                catch ( MsalUiRequiredException )
                {
                    interactionRequired = true;
                }
                catch ( Exception ex )
                {
                    Console.WriteLine( $"Authentication error: {ex.Message}" );
                }

                if ( interactionRequired )
                {
                    try
                    {
                        result = await Pca.AcquireTokenInteractive( Scopes ).ExecuteAsync();
                    } catch (Exception ex )
                    {
                        Console.Write( $"Authentication error: {ex.Message}" );
                    }
                }

                Console.WriteLine( $"Access Token: {result.AccessToken}\n" );
                Console.WriteLine( $"Graph Request: {request.RequestUri}" );

                // Set the access token for the current request
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue( "Bearer", result.AccessToken );

            }
        }
    }
}
