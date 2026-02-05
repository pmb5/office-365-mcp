using System.Net;
using Microsoft.Extensions.Logging;
using NSubstitute;
using Office365MCP.Auth;
using Office365MCP.Graph;
using Xunit;

namespace Office365MCP.Tests.Graph;

public class GraphClientServiceTests
{
    private readonly HttpClient _httpClient;
    private readonly IGraphAuthService _authService;
    private readonly ILogger<GraphClientService> _logger;
    private readonly GraphClientService _service;

    public GraphClientServiceTests()
    {
        _httpClient = new HttpClient
        {
            BaseAddress = new Uri("https://graph.microsoft.com/v1.0/")
        };
        _authService = Substitute.For<IGraphAuthService>();
        _logger = Substitute.For<ILogger<GraphClientService>>();

        _service = new GraphClientService(_httpClient, _authService, _logger);
    }

    [Fact]
    public void Constructor_WithValidParameters_CreatesInstance()
    {
        Assert.NotNull(_service);
    }

    [Theory]
    [InlineData("/me/messages", null)]
    [InlineData("/me/messages", "")]
    public void BuildUrl_WithNoQueryParams_ReturnsEndpoint(string endpoint, string? ignored)
    {
        var queryParams = ignored != null ? new Dictionary<string, string>() : null;
        var result = typeof(GraphClientService)
            .GetMethod("BuildUrl", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)?
            .Invoke(null, new object?[] { endpoint, queryParams }) as string;

        Assert.Equal(endpoint, result);
    }

    [Fact]
    public void BuildUrl_WithQueryParams_ReturnsFormattedUrl()
    {
        var endpoint = "/me/messages";
        var queryParams = new Dictionary<string, string>
        {
            ["$top"] = "10",
            ["$select"] = "id,subject"
        };

        var result = typeof(GraphClientService)
            .GetMethod("BuildUrl", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)?
            .Invoke(null, new object?[] { endpoint, queryParams }) as string;

        Assert.Contains("%24top=10", result); // $top is escaped to %24top
        Assert.Contains("%24select=id%2Csubject", result); // $select and comma are escaped
    }
}
