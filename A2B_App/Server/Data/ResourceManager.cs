using IdentityModel;
using IdentityServer4.Models;
using System.Collections.Generic;
using IdentityServer4;
//using static IdentityServer4.IdentityServerConstants;

namespace A2B_App.Server.Data
{
    internal static class ResourceManager
    {
        //public static IEnumerable<ApiResource> Apis =>
        //    new List<ApiResource>
        //    {
        //        new ApiResource {
        //            Name = "app.api.whatever",
        //            DisplayName = "Whatever Apis",
        //            ApiSecrets = { new Secret("a75a559d-1dab-4c65-9bc0-f8e590cb388d".Sha256()) },
        //            Scopes = new List<Scope> {
        //                new Scope("app.api.whatever.read"),
        //                new Scope("app.api.whatever.write"),
        //                new Scope("app.api.whatever.full")
        //            }
        //        },
        //        new ApiResource("app.api.weather","Weather Apis"),
        //        // local API
        //        new ApiResource(LocalApi.ScopeName),

        //        // simple version with ctor
        //        new ApiResource("api1", "A2B_App.ServerAPI")
        //        {
        //            // this is needed for introspection when using reference tokens
        //            ApiSecrets = { new Secret("secret".Sha256()) }
        //        },
        //    };

        public static IEnumerable<ApiResource> GetApiResources()
        {
            return new List<ApiResource>
            {
                new ApiResource("A2B_App.ServerAPI", "A2B_App.ServerAPI"),
                //new ApiResource("A2B_App.ServerAPI", "A2B_App.ServerAPI"),
                new ApiResource {
                    Name = "app.api.whatever",
                    DisplayName = "Whatever Apis",
                    ApiSecrets = { new Secret("a75a559d-1dab-4c65-9bc0-f8e590cb388d".Sha256()) },
                    Scopes = new List<Scope> {
                        new Scope("app.api.whatever.read"),
                        new Scope("app.api.whatever.write"),
                        new Scope("app.api.whatever.full")
                    }
                },
                new ApiResource("app.api.weather","Weather Apis")
            };
        }

        public static List<IdentityResource> GetIdentityResources()
        {
            //// Claims automatically included in OpenId scope
            //var openIdScope = new IdentityResources.OpenId();
            //openIdScope.UserClaims.Add(JwtClaimTypes.Locale);

            return new List<IdentityResource>
            {

                new IdentityResources.OpenId(),
                new IdentityResources.Profile(),
                //new IdentityResources.Address(),
                new IdentityResources.Email(),
                //new IdentityResources.Phone(),               

            };
        }

    }
}
