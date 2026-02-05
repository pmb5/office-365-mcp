namespace Office365MCP.Auth;

public interface IGraphAuthService
{
    Task<string> GetAccessTokenAsync();
    Task ClearTokenCacheAsync();
}
