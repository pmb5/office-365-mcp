using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Extensions.Msal;

namespace Office365MCP.Auth;

public class GraphAuthService : IGraphAuthService
{
    private readonly IPublicClientApplication _app;
    private readonly string[] _scopes;
    private readonly ILogger<GraphAuthService> _logger;
    private MsalCacheHelper? _cacheHelper;

    public GraphAuthService(IConfiguration configuration, ILogger<GraphAuthService> logger)
    {
        _logger = logger;

        var clientId = configuration["Graph:ClientId"]
            ?? throw new InvalidOperationException("Graph:ClientId not configured");
        var tenantId = configuration["Graph:TenantId"]
            ?? throw new InvalidOperationException("Graph:TenantId not configured");

        _scopes = configuration.GetSection("Graph:Scopes").Get<string[]>()
            ?? throw new InvalidOperationException("Graph:Scopes not configured");

        var authority = $"https://login.microsoftonline.com/{tenantId}";

        _app = PublicClientApplicationBuilder.Create(clientId)
            .WithAuthority(authority)
            .WithDefaultRedirectUri()
            .Build();

        InitializeCacheAsync().Wait();
    }

    private async Task InitializeCacheAsync()
    {
        try
        {
            var storageProperties = new StorageCreationPropertiesBuilder(
                "office365mcp.cache",
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Office365MCP"))
                .Build();

            _cacheHelper = await MsalCacheHelper.CreateAsync(storageProperties);
            _cacheHelper.RegisterCache(_app.UserTokenCache);

            _logger.LogInformation("Token cache initialized successfully");
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to initialize token cache, continuing without persistent cache");
        }
    }

    public async Task<string> GetAccessTokenAsync()
    {
        try
        {
            var accounts = await _app.GetAccountsAsync();
            var firstAccount = accounts.FirstOrDefault();

            if (firstAccount != null)
            {
                try
                {
                    var result = await _app.AcquireTokenSilent(_scopes, firstAccount)
                        .ExecuteAsync();

                    _logger.LogInformation("Token acquired silently");
                    return result.AccessToken;
                }
                catch (MsalUiRequiredException)
                {
                    _logger.LogInformation("Silent token acquisition failed, falling back to interactive");
                }
            }

            var interactiveResult = await _app.AcquireTokenInteractive(_scopes)
                .WithPrompt(Prompt.SelectAccount)
                .ExecuteAsync();

            _logger.LogInformation("Token acquired interactively");
            return interactiveResult.AccessToken;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to acquire access token");
            throw;
        }
    }

    public async Task ClearTokenCacheAsync()
    {
        try
        {
            var accounts = await _app.GetAccountsAsync();
            foreach (var account in accounts)
            {
                await _app.RemoveAsync(account);
            }

            _logger.LogInformation("Token cache cleared");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to clear token cache");
            throw;
        }
    }
}
