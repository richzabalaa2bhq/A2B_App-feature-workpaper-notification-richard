using System;
using System.Net.Http;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Text;
using Microsoft.AspNetCore.Components.WebAssembly.Authentication;
using Microsoft.AspNetCore.Components.WebAssembly.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using A2B_App.Shared;
using Blazored.Modal;
using Blazored.Toast;
using Microsoft.AspNetCore.Components;
using System.Net.Http.Json;

namespace A2B_App.Client
{
    public class Program
    {
        public static async Task Main(string[] args)
        {
            var builder = WebAssemblyHostBuilder.CreateDefault(args);
            builder.RootComponents.Add<App>("app");

            builder.Services.AddHttpClient("A2B_App.ServerAPI", client => { 
                client.BaseAddress = new Uri(builder.HostEnvironment.BaseAddress); 
                client.Timeout = TimeSpan.FromMinutes(30); 
            }).AddHttpMessageHandler<BaseAddressAuthorizationMessageHandler>();

            //builder.Services.AddHttpClient("A2B_App.ServerAPI",
            //    client => client.BaseAddress = new Uri("https://ddohq.com/base"))
            //.AddHttpMessageHandler(sp => sp.GetRequiredService<AuthorizationMessageHandler>()
            //.ConfigureHandler(
            //    authorizedUrls: new[] { "https://ddohq.com/base" },
            //    scopes: new[] { "ddohq.read", "ddohq.write" }));

            builder.Services.AddBlazoredToast();
            builder.Services.AddBlazoredModal();

            // Supply HttpClient instances that include access tokens when making requests to the server project
            builder.Services.AddTransient(sp => sp.GetRequiredService<IHttpClientFactory>().CreateClient("A2B_App.ServerAPI"));

            //builder.Services.AddApiAuthorization();
            builder.Services.AddApiAuthorization()
                .AddAccountClaimsPrincipalFactory<RolesClaimsPrincipalFactory>();

            builder.Services.AddSingleton(async p =>
            {
                var httpClient = p.GetRequiredService<HttpClient>();
                return await httpClient.GetFromJsonAsync<ClientSettings>("settings.json").ConfigureAwait(false);
                //.GetJsonAsync<ClientSettings>("settings.json").ConfigureAwait(false);
            });

            await builder.Build().RunAsync();
        }
    }
}
