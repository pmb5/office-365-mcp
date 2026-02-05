namespace Office365MCP.Graph;

public interface IGraphClientService
{
    Task<string> GetAsync(string endpoint, Dictionary<string, string>? queryParams = null);
    Task<string> PostAsync(string endpoint, string jsonBody);
}
