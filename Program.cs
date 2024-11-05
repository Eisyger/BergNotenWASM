using BergNotenWASM;
using BergNotenWASM.Services;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.Identity.Web;
using Microsoft.AspNetCore.Components.WebAssembly.Hosting;
using Microsoft.AspNetCore.Components.WebAssembly.Authentication;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using System.Net.Http.Headers;

var builder = WebAssemblyHostBuilder.CreateDefault(args);
builder.RootComponents.Add<App>("#app");
builder.RootComponents.Add<HeadOutlet>("head::after");

#region Auth
builder.Configuration.AddJsonFile("appsettings.json");

builder.Services.AddMsalAuthentication(options =>
{
    var azureOptions = builder.Configuration.GetSection("Azure");
    options.ProviderOptions.DefaultAccessTokenScopes.Add($"{azureOptions["ApiId"]}/.default");
    options.ProviderOptions.Authority = $"https://login.microsoftonline.com/{azureOptions["TenantId"]}/";
    options.ProviderOptions.ClientId = azureOptions["ClientId"];
});
#endregion 

builder.Services.AddScoped(_ => new HttpClient { BaseAddress = new Uri(builder.HostEnvironment.BaseAddress) });

// Registriere den LocalStorageService
builder.Services.AddScoped<LocalStorageService>();

await builder.Build().RunAsync();
