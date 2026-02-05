using System.Net;
using System.Text;
using Microsoft.Extensions.Logging;
using Office365MCP.Auth;

namespace Office365MCP.Graph;

public class GraphClientService : IGraphClientService
{
    private readonly HttpClient _httpClient;
    private readonly IGraphAuthService _authService;
    private readonly ILogger<GraphClientService> _logger;

    public GraphClientService(
        HttpClient httpClient,
        IGraphAuthService authService,
        ILogger<GraphClientService> logger)
    {
        _httpClient = httpClient;
        _authService = authService;
        _logger = logger;
    }

    public async Task<string> GetAsync(string endpoint, Dictionary<string, string>? queryParams = null)
    {
        var url = BuildUrl(endpoint, queryParams);
        return await ExecuteWithRetryAsync(async () =>
        {
            var request = new HttpRequestMessage(HttpMethod.Get, url);
            return await SendRequestAsync(request);
        });
    }

    public async Task<string> PostAsync(string endpoint, string jsonBody)
    {
        return await ExecuteWithRetryAsync(async () =>
        {
            var request = new HttpRequestMessage(HttpMethod.Post, endpoint)
            {
                Content = new StringContent(jsonBody, Encoding.UTF8, "application/json")
            };
            return await SendRequestAsync(request);
        });
    }

    private async Task<string> ExecuteWithRetryAsync(Func<Task<string>> action)
    {
        try
        {
            return await action();
        }
        catch (UnauthorizedException)
        {
            _logger.LogWarning("Received 401 Unauthorized, clearing cache and retrying");
            await _authService.ClearTokenCacheAsync();
            return await action();
        }
    }

    private async Task<string> SendRequestAsync(HttpRequestMessage request)
    {
        var token = await _authService.GetAccessTokenAsync();
        request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

        var response = await _httpClient.SendAsync(request);

        if (response.StatusCode == HttpStatusCode.Unauthorized)
        {
            throw new UnauthorizedException();
        }

        response.EnsureSuccessStatusCode();
        return await response.Content.ReadAsStringAsync();
    }

    private static string BuildUrl(string endpoint, Dictionary<string, string>? queryParams)
    {
        if (queryParams == null || queryParams.Count == 0)
        {
            return endpoint;
        }

        var queryString = string.Join("&",
            queryParams.Select(kvp => $"{Uri.EscapeDataString(kvp.Key)}={Uri.EscapeDataString(kvp.Value)}"));

        return $"{endpoint}?{queryString}";
    }

    private class UnauthorizedException : Exception
    {
    }
}
