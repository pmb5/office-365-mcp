using Microsoft.Extensions.Logging;
using NSubstitute;
using Office365MCP.Graph;
using Office365MCP.Tools;
using Office365MCP.Tools.Search;
using System.Text.Json;
using Xunit;

namespace Office365MCP.Tests.Tools;

public class OutlookMailToolsTests
{
    private readonly IGraphClientService _graphClient;
    private readonly ILogger<OutlookMailTools> _logger;
    private readonly OutlookMailTools _tools;

    public OutlookMailToolsTests()
    {
        _graphClient = Substitute.For<IGraphClientService>();
        _logger = Substitute.For<ILogger<OutlookMailTools>>();
        _tools = new OutlookMailTools(_graphClient, _logger);
    }

    [Fact]
    public void Constructor_WithValidParameters_CreatesInstance()
    {
        Assert.NotNull(_tools);
    }

    [Fact]
    public async Task ListEmails_WithDefaultCount_CallsGraphClientWithCorrectParams()
    {
        var expectedResponse = "{\"value\":[]}";
        _graphClient.GetAsync(Arg.Any<string>(), Arg.Any<Dictionary<string, string>>())
            .Returns(expectedResponse);

        var result = await _tools.ListEmails();

        Assert.Equal(expectedResponse, result);
        await _graphClient.Received(1).GetAsync(
            "me/messages",
            Arg.Is<Dictionary<string, string>>(d =>
                d["$top"] == "10" &&
                d.ContainsKey("$select") &&
                d["$orderby"] == "receivedDateTime desc"));
    }

    [Fact]
    public async Task ListEmails_WithCustomCount_CallsGraphClientWithCorrectCount()
    {
        var expectedResponse = "{\"value\":[]}";
        _graphClient.GetAsync(Arg.Any<string>(), Arg.Any<Dictionary<string, string>>())
            .Returns(expectedResponse);

        var result = await _tools.ListEmails(25);

        Assert.Equal(expectedResponse, result);
        await _graphClient.Received(1).GetAsync(
            "me/messages",
            Arg.Is<Dictionary<string, string>>(d => d["$top"] == "25"));
    }

    [Fact]
    public async Task GetEmail_WithValidMessageId_CallsGraphClient()
    {
        var messageId = "test-message-id";
        var expectedResponse = "{\"id\":\"test-message-id\"}";
        _graphClient.GetAsync(Arg.Any<string>(), Arg.Any<Dictionary<string, string>>())
            .Returns(expectedResponse);

        var result = await _tools.GetEmail(messageId);

        Assert.Equal(expectedResponse, result);
        await _graphClient.Received(1).GetAsync($"me/messages/{messageId}", null);
    }

    [Fact]
    public async Task SearchEmails_WithQuery_CallsGraphClientWithSearchParam()
    {
        var query = "from:john@example.com";
        _graphClient.GetAsync(Arg.Any<string>(), Arg.Any<Dictionary<string, string>>())
            .Returns("{\"value\":[]}");

        await _tools.SearchEmails(query);

        await _graphClient.Received(1).GetAsync(
            "me/messages",
            Arg.Is<Dictionary<string, string>>(d =>
                d["$search"] == $"\"{query}\"" &&
                d["$top"] == "10"));
    }

    [Fact]
    public async Task SearchEmails_DoesNotSendOrderBy()
    {
        _graphClient.GetAsync(Arg.Any<string>(), Arg.Any<Dictionary<string, string>>())
            .Returns("{\"value\":[]}");

        await _tools.SearchEmails("receipt");

        await _graphClient.Received(1).GetAsync(
            "me/messages",
            Arg.Is<Dictionary<string, string>>(d => !d.ContainsKey("$orderby")));
    }

    [Fact]
    public async Task SearchEmails_WithQuotedPhrase_EscapesInnerQuotes()
    {
        var query = "subject:\"Your receipt from Anthropic, PBC\"";
        _graphClient.GetAsync(Arg.Any<string>(), Arg.Any<Dictionary<string, string>>())
            .Returns("{\"value\":[]}");

        await _tools.SearchEmails(query);

        await _graphClient.Received(1).GetAsync(
            "me/messages",
            Arg.Is<Dictionary<string, string>>(d =>
                d["$search"] == "\"subject:\\\"Your receipt from Anthropic, PBC\\\"\""));
    }

    [Fact]
    public async Task SearchEmails_WithFilters_CombinesThemIntoOneSearch()
    {
        _graphClient.GetAsync(Arg.Any<string>(), Arg.Any<Dictionary<string, string>>())
            .Returns("{\"value\":[]}");

        await _tools.SearchEmails(from: "billing@example.com", startDate: "2026-09-01", hasAttachments: true);

        await _graphClient.Received(1).GetAsync(
            "me/messages",
            Arg.Is<Dictionary<string, string>>(d =>
                d["$search"] == "\"from:billing@example.com AND received>=2026-09-01 AND hasattachments:true\""));
    }

    [Fact]
    public async Task SearchEmails_WithNextLink_ReturnsTokenThatFollowsIt()
    {
        const string nextLink = "https://graph.microsoft.com/v1.0/me/messages?$search=x&$skiptoken=abc";
        _graphClient.GetAsync("me/messages", Arg.Any<Dictionary<string, string>>())
            .Returns(JsonSerializer.Serialize(new Dictionary<string, object> { ["value"] = Array.Empty<object>(), ["@odata.nextLink"] = nextLink }));
        _graphClient.GetAsync(nextLink, null).Returns("{\"value\":[]}");

        var first = JsonSerializer.Deserialize<JsonElement>(await _tools.SearchEmails("x"));
        await _tools.SearchEmails(nextPage: first.GetProperty("nextPage").GetString());

        await _graphClient.Received(1).GetAsync(nextLink, null);
    }

    [Fact]
    public async Task SearchEmails_WithNextPageForAnotherHost_DoesNotCallGraph()
    {
        var token = SearchCursor.Encode("https://attacker.example.com/me/messages?x=1");

        var result = await _tools.SearchEmails(nextPage: token);

        Assert.Contains("error", result);
        await _graphClient.DidNotReceive().GetAsync(Arg.Any<string>(), Arg.Any<Dictionary<string, string>?>());
    }

    [Fact]
    public async Task SearchEmails_PreviewWithInvisiblePadding_IsCollapsed()
    {
        var message = new { id = "1", bodyPreview = "Your receipt  ͏  ͏  ͏ end" };
        _graphClient.GetAsync(Arg.Any<string>(), Arg.Any<Dictionary<string, string>>())
            .Returns(JsonSerializer.Serialize(new { value = new[] { message } }));

        var result = JsonSerializer.Deserialize<JsonElement>(await _tools.SearchEmails("receipt"));

        Assert.Equal("Your receipt end", result.GetProperty("results")[0].GetProperty("preview").GetString());
    }

    [Fact]
    public async Task SendEmail_WithValidParameters_CallsGraphClientPost()
    {
        var to = "recipient@example.com";
        var subject = "Test Subject";
        var body = "Test Body";

        _graphClient.PostAsync(Arg.Any<string>(), Arg.Any<string>())
            .Returns("{}");

        var result = await _tools.SendEmail(to, subject, body);

        Assert.Contains("success", result);
        await _graphClient.Received(1).PostAsync(
            "me/sendMail",
            Arg.Is<string>(json =>
                json.Contains(to) &&
                json.Contains(subject) &&
                json.Contains(body)));
    }

    [Fact]
    public async Task ListMailFolders_CallsGraphClient()
    {
        var expectedResponse = "{\"value\":[]}";
        _graphClient.GetAsync(Arg.Any<string>(), Arg.Any<Dictionary<string, string>>())
            .Returns(expectedResponse);

        var result = await _tools.ListMailFolders();

        Assert.Equal(expectedResponse, result);
        await _graphClient.Received(1).GetAsync("me/mailFolders", null);
    }
}
