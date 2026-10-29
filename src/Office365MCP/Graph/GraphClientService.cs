using System.Net;
using System.Text;
using Microsoft.Extensions.Logging;
using Office365MCP.Auth;

namespace Office365MCP.Graph;

public class GraphClientService : IGraphClientService
{
    private readonly HttpClient _httpClient;
    private readonly IGraphAuthService _authService;
    private readonly IPermissionHelper _permissionHelper;
    private readonly ILogger<GraphClientService> _logger;

    public GraphClientService(
        HttpClient httpClient,
        IGraphAuthService authService,
        IPermissionHelper permissionHelper,
        ILogger<GraphClientService> logger)
    {
        _httpClient = httpClient;
        _authService = authService;
        _permissionHelper = permissionHelper;
        _logger = logger;
    }

    public async Task<string> GetAsync(string endpoint, Dictionary<string, string>? queryParams = null)
    {
        var url = BuildUrl(endpoint, queryParams);
        return await ExecuteWithRetryAsync(async () =>
        {
            var request = new HttpRequestMessage(HttpMethod.Get, url);
            if (RequiresEventualConsistency(endpoint, queryParams))
            {
                request.Headers.Add("ConsistencyLevel", "eventual");
            }
            return await SendRequestAsync(request);
        });
    }

    public Task<string> PostAsync(string endpoint, string jsonBody)
    {
        return PostContentAsync(endpoint, jsonBody, "application/json");
    }

    public Task<string> PostHtmlAsync(string endpoint, string html)
    {
        return PostContentAsync(endpoint, html, "text/html");
    }

    private async Task<string> PostContentAsync(string endpoint, string body, string mediaType)
    {
        return await ExecuteWithRetryAsync(async () =>
        {
            var request = new HttpRequestMessage(HttpMethod.Post, endpoint)
            {
                Content = new StringContent(body, Encoding.UTF8, mediaType)
            };
            return await SendRequestAsync(request);
        });
    }

    // Directory objects only accept $search with this header; other resources reject or ignore it.
    private static bool RequiresEventualConsistency(string endpoint, Dictionary<string, string>? queryParams)
    {
        return queryParams != null
            && queryParams.ContainsKey("$search")
            && DirectoryResources.Any(r => endpoint.TrimStart('/').StartsWith(r, StringComparison.OrdinalIgnoreCase));
    }

    private static readonly string[] DirectoryResources = { "users", "groups", "devices", "applications", "servicePrincipals" };

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
        catch (InsufficientPermissionException ex)
        {
            _logger.LogWarning("Insufficient permissions detected");

            // Prompt user to add the missing permission
            if (await _permissionHelper.HandlePermissionErrorAsync(ex.ErrorResponse))
            {
                _logger.LogInformation("Permission added, retrying operation");
                // The permission was added and token cache cleared, retry the operation
                return await action();
            }

            // User declined or permission couldn't be added, throw the original exception
            throw;
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

        if (response.StatusCode == HttpStatusCode.Forbidden)
        {
            var errorContent = await response.Content.ReadAsStringAsync();

            // Check if it's a permission error
            if (errorContent.Contains("Insufficient privileges", StringComparison.OrdinalIgnoreCase) ||
                errorContent.Contains("Authorization_RequestDenied", StringComparison.OrdinalIgnoreCase) ||
                errorContent.Contains("Access is denied", StringComparison.OrdinalIgnoreCase) ||
                errorContent.Contains("does not have permission", StringComparison.OrdinalIgnoreCase))
            {
                throw new InsufficientPermissionException(errorContent);
            }
        }

        var content = await response.Content.ReadAsStringAsync();
        if (!response.IsSuccessStatusCode)
        {
            // Graph explains a rejected request in the body; without it a bad query only shows as "400".
            throw new HttpRequestException(
                $"Graph returned {(int)response.StatusCode} ({response.ReasonPhrase}): {DescribeGraphError(content)}",
                null,
                response.StatusCode);
        }
        return content;
    }

    private static string DescribeGraphError(string content)
    {
        try
        {
            using var doc = System.Text.Json.JsonDocument.Parse(content);
            if (doc.RootElement.TryGetProperty("error", out var error))
            {
                var code = error.TryGetProperty("code", out var c) ? c.GetString() : null;
                var message = error.TryGetProperty("message", out var m) ? m.GetString() : null;
                return string.IsNullOrEmpty(code) ? message ?? content : $"{code}: {message}";
            }
        }
        catch (System.Text.Json.JsonException)
        {
        }
        return string.IsNullOrWhiteSpace(content) ? "no details" : content;
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

    private class InsufficientPermissionException : Exception
    {
        public string ErrorResponse { get; }

        public InsufficientPermissionException(string errorResponse)
            : base("Insufficient permissions")
        {
            ErrorResponse = errorResponse;
        }
    }
}
