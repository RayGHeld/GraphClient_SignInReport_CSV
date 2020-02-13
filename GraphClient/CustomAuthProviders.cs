using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Identity.Client;

namespace GraphClient
{
    /// <summary>
    /// This is a custom set of classes to implement the IAuthenticationProvider class.
    /// </summary>
    class CustomAuthProviders
    {
        /// <summary>
        /// Prompts the user to sign-in with a dialog
        /// </summary>
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
                IEnumerable<IAccount> accounts;
                AuthenticationResult result = null;
                bool interactionRequired = false;

                accounts = await Pca.GetAccountsAsync();

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
                    Console.WriteLine( $"Authentication error: { ex.Message }" );
                    return;
                }

                if ( interactionRequired )
                {
                    try
                    {
                        result = await Pca.AcquireTokenInteractive( Scopes ).ExecuteAsync();
                    }
                    catch ( Exception ex )
                    {
                        Console.WriteLine( $"Authentication error: { ex.Message }" );
                        return;
                    }
                }

                Console.WriteLine( $"Access Token: { result.AccessToken }{ Environment.NewLine }" );
                Console.WriteLine( $"Graph Request: { request.RequestUri }" );

                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue( "Bearer", result.AccessToken );

            }

        }

        /// <summary>
        /// Implements the client credentials flow ( no user context ) : https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki/Client-credential-flows
        /// </summary>
        public class ClientCredentialsProvider : IAuthenticationProvider
        {
            private static IConfidentialClientApplication Cca = null;
            private static List<String> Scopes = null;

            private ClientCredentialsProvider()
            {
                // intentionally left blank and private to prevent empty constructor
            }

            public ClientCredentialsProvider( IConfidentialClientApplication cca, List<String> scopes )
            {
                Cca = cca;
                Scopes = scopes;
            }

            public async Task AuthenticateRequestAsync( HttpRequestMessage request )
            {
                AuthenticationResult result = null;

                try
                {
                    result = await Cca.AcquireTokenForClient( Scopes ).ExecuteAsync();
                } catch (MsalServiceException ex )
                {
                    // case when ex.message contains: 
                    // AADSTS7011 Invalid scope. The scope has to be in the form "https://graph.microsoft.com/.default"

                    Console.WriteLine( $"Invalid scope parameter: { ex.Message }" );
                    return;
                }

                Console.WriteLine( $"Access Token: { result.AccessToken }{ Environment.NewLine }" );
                Console.WriteLine( $"Graph Request: { request.RequestUri }" );

                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue( "Bearer", result.AccessToken );

            }
        }
    }
}
