using IdentityServer4;
using IdentityServer4.Models;
using System.Collections.Generic;

namespace A2B_App.Server.Data
{
    internal static class ClientManager
    {
        public static IEnumerable<IdentityServer4.Models.Client> Clients =>
            new List<IdentityServer4.Models.Client>
            {
                    new IdentityServer4.Models.Client
                    {
                         ClientName = "Client Application1",
                         ClientId = "t8agr5xKt4$3",
                         AllowedGrantTypes = GrantTypes.ClientCredentials,
                         ClientSecrets = { new Secret("eb300de4-add9-42f4-a3ac-abd3c60f1919".Sha256()) },
                         AllowedScopes = new List<string> { "app.api.whatever.read", "app.api.whatever.write" },
                         Enabled = true
                    },
                    new IdentityServer4.Models.Client
                    {
                         ClientName = "device",
                         ClientId = "device",
                         AllowedGrantTypes = GrantTypes.ClientCredentials,
                         ClientSecrets = { new Secret("1554db43-3015-47a8-a748-55bd76b6af48".Sha256()) },
                         AllowedScopes = { "app.api.weather" },
                         Enabled = true
                    },
                    new IdentityServer4.Models.Client
                    {
                        ClientId = "A2B_App.Client",
                        ClientName = "A2B_App.Client",
                        AllowedGrantTypes = GrantTypes.CodeAndClientCredentials,
                        RequirePkce = true,
                        RequireClientSecret = false,
                        AllowedCorsOrigins = { "https://localhost:5001" },
                        RedirectUris = { "https://localhost:5001/authentication/login-callback" },
                        PostLogoutRedirectUris = { "https://localhost:5001/" },
                        Enabled = true,

                        AllowedScopes =
                        {
                            IdentityServerConstants.StandardScopes.OpenId,
                            IdentityServerConstants.StandardScopes.Profile,
                            "A2B_App.ServerAPI"
                        },

                    }
            };
    }
}
