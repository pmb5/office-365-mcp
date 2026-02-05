using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Server;
using Office365MCP.Auth;
using Office365MCP.Graph;
using Office365MCP.Tools;

var builder = Host.CreateApplicationBuilder(args);

builder.Configuration.AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

builder.Logging.ClearProviders()
    .AddConsole(options => options.LogToStandardErrorThreshold = LogLevel.Trace);

builder.Services.AddSingleton<IScopeProvider, ScopeProvider>();
builder.Services.AddSingleton<IGraphAuthService, GraphAuthService>();
builder.Services.AddSingleton<IPermissionHelper, PermissionHelper>();
builder.Services.AddHttpClient<IGraphClientService, GraphClientService>(client =>
{
    client.BaseAddress = new Uri("https://graph.microsoft.com/v1.0/");
});

builder.Services.AddSingleton<OutlookMailTools>();
builder.Services.AddSingleton<OutlookCalendarTools>();
builder.Services.AddSingleton<TeamsTools>();
builder.Services.AddSingleton<OneNoteTools>();
builder.Services.AddSingleton<PermissionManagementTools>();

builder.Services
    .AddMcpServer()
    .WithStdioServerTransport()
    .WithToolsFromAssembly(typeof(Program).Assembly);

var host = builder.Build();
await host.RunAsync();
