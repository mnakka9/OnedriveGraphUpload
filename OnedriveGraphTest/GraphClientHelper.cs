using Microsoft.Identity.Client;
using System;
using System.Linq;
using System.Threading;
using System.Web;
using Microsoft.Owin.Security;
using Microsoft.Owin.Security.OpenIdConnect;
using System.Security.Claims;
using System.Threading.Tasks;
using Microsoft.Graph;
using System.Net.Http.Headers;
using OnedriveGraphTest;
using System.Diagnostics;

namespace OnedriveGraphTest
{
    public static class GraphClientHelper
    {
        private static string ClientId = "7c70606c-c585-4850-aa15-5717ee1227b9";
        private static String[] Scopes = {
                        "User.Read",
                        "User.ReadBasic.All",
                        "Mail.Send",
                        "Mail.Read",
                        "Group.ReadWrite.All",
                        "Sites.Read.All",
                        "Directory.AccessAsUser.All",
                        "Files.ReadWrite",
                        "Files.ReadWrite.AppFolder"
                        };
        public static PublicClientApplication IdentityClientApp = new PublicClientApplication(ClientId);

        public static string TokenForUser = null;
        public static DateTimeOffset Expiration;

        private static GraphServiceClient graphClient = null;

        // Get an access token for the given context and resourceId. An attempt is first made to 
        // acquire the token silently. If that fails, then we try to acquire the token by prompting the user.
        public static GraphServiceClient GetAuthenticatedClient()
        {
            if (graphClient == null)
            {
                // Create Microsoft Graph client.
                try
                {
                    graphClient = new GraphServiceClient(
                        "https://graph.microsoft.com/v1.0",
                        new DelegateAuthenticationProvider(
                            async (requestMessage) =>
                            {
                                var token = await GetTokenForUserAsync(IdentityClientApp);
                                requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);

                            }));
                    return graphClient;
                }

                catch (Exception ex)
                {
                    Debug.WriteLine("Could not create a graph client: " + ex.Message);
                }
            }

            return graphClient;
        }


        /// <summary>
        /// Get Token for User.
        /// </summary>
        /// <returns>Token for user.</returns>
        public static async Task<string> GetTokenForUserAsync(PublicClientApplication clientApp)
        {
            AuthenticationResult authResult;
            try
            {
                if (clientApp.Users.Count() > 0)
                {
                    authResult = await IdentityClientApp.AcquireTokenSilentAsync(Scopes,clientApp.Users.FirstOrDefault());
                    TokenForUser = authResult.AccessToken;
                }
                else
                {
                    throw new Exception("Get correct access token");
                }
                
            }

            catch (Exception)
            {
                if (TokenForUser == null || Expiration <= DateTimeOffset.UtcNow.AddMinutes(5))
                {
                    authResult = await IdentityClientApp.AcquireTokenAsync(Scopes);

                    TokenForUser = authResult.AccessToken;
                    Expiration = authResult.ExpiresOn;
                }
            }

            return TokenForUser;
        }

        /// <summary>
        /// Signs the user out of the service.
        /// </summary>
        public static void SignOut()
        {
            graphClient = null;
            TokenForUser = null;

        }
    }

}