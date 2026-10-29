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
    private readonly IPermissionHelper _permissionHelper;
    private readonly ILogger<GraphClientService> _logger;
    private readonly GraphClientService _service;

    public GraphClientServiceTests()
    {
        _httpClient = new HttpClient
        {
            BaseAddress = new Uri("https://graph.microsoft.com/v1.0/")
        };
        _authService = Substitute.For<IGraphAuthService>();
        _permissionHelper = Substitute.For<IPermissionHelper>();
        _logger = Substitute.For<ILogger<GraphClientService>>();

        _service = new GraphClientService(_httpClient, _authService, _permissionHelper, _logger);
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

    private sealed class StubHandler : HttpMessageHandler
    {
        private readonly HttpResponseMessage _response;
        public HttpRequestMessage? LastRequest { get; private set; }

        public StubHandler(HttpResponseMessage response) => _response = response;

        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            LastRequest = request;
            return Task.FromResult(_response);
        }
    }

    private GraphClientService CreateService(StubHandler handler)
    {
        var client = new HttpClient(handler) { BaseAddress = new Uri("https://graph.microsoft.com/v1.0/") };
        _authService.GetAccessTokenAsync().Returns("token");
        return new GraphClientService(client, _authService, _permissionHelper, _logger);
    }

    [Fact]
    public async Task GetAsync_OnBadRequest_IncludesGraphErrorInMessage()
    {
        var handler = new StubHandler(new HttpResponseMessage(HttpStatusCode.BadRequest)
        {
            Content = new StringContent("{\"error\":{\"code\":\"BadRequest\",\"message\":\"$orderby is not supported with $search\"}}")
        });

        var ex = await Assert.ThrowsAsync<HttpRequestException>(() => CreateService(handler).GetAsync("me/messages"));

        Assert.Equal(HttpStatusCode.BadRequest, ex.StatusCode);
        Assert.Contains("BadRequest: $orderby is not supported with $search", ex.Message);
    }

    [Fact]
    public async Task GetAsync_SearchOnUsers_SendsConsistencyLevelHeader()
    {
        var handler = new StubHandler(new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent("{}") });

        await CreateService(handler).GetAsync("users", new Dictionary<string, string> { ["$search"] = "\"displayName:x\"" });

        Assert.Equal("eventual", handler.LastRequest!.Headers.GetValues("ConsistencyLevel").Single());
    }

    [Fact]
    public async Task GetAsync_SearchOnMessages_DoesNotSendConsistencyLevelHeader()
    {
        var handler = new StubHandler(new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent("{}") });

        await CreateService(handler).GetAsync("me/messages", new Dictionary<string, string> { ["$search"] = "\"x\"" });

        Assert.False(handler.LastRequest!.Headers.Contains("ConsistencyLevel"));
    }

    [Fact]
    public async Task PostHtmlAsync_SendsTextHtmlContentType()
    {
        var handler = new StubHandler(new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent("{}") });

        await CreateService(handler).PostHtmlAsync("me/onenote/sections/s/pages", "<html></html>");

        Assert.Equal("text/html", handler.LastRequest!.Content!.Headers.ContentType!.MediaType);
    }
}
