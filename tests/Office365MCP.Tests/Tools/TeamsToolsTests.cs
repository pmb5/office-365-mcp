using Microsoft.Extensions.Logging;
using NSubstitute;
using Office365MCP.Graph;
using Office365MCP.Tools;
using System.Text.Json;
using Xunit;

namespace Office365MCP.Tests.Tools;

public class TeamsToolsTests
{
    private readonly IGraphClientService _graphClient;
    private readonly TeamsTools _tools;

    public TeamsToolsTests()
    {
        _graphClient = Substitute.For<IGraphClientService>();
        _tools = new TeamsTools(_graphClient, Substitute.For<ILogger<TeamsTools>>());
    }

    private static string Page(string? nextLink, params (string body, string sender)[] messages)
    {
        var value = messages.Select(m => new
        {
            body = new { content = m.body },
            from = new { user = new { displayName = m.sender } }
        });
        return nextLink == null
            ? JsonSerializer.Serialize(new { value })
            : JsonSerializer.Serialize(new Dictionary<string, object> { ["value"] = value, ["@odata.nextLink"] = nextLink });
    }

    private static int ResultCount(string json) =>
        JsonSerializer.Deserialize<JsonElement>(json).GetProperty("value").GetArrayLength();

    [Fact]
    public async Task SearchChatMessages_DoesNotSendSearchParam()
    {
        _graphClient.GetAsync(Arg.Any<string>(), Arg.Any<Dictionary<string, string>?>())
            .Returns(Page(null));

        await _tools.SearchChatMessages("chat1", "budget");

        await _graphClient.Received(1).GetAsync(
            "chats/chat1/messages",
            Arg.Is<Dictionary<string, string>>(d => !d.ContainsKey("$search") && d["$top"] == "50"));
    }

    [Fact]
    public async Task SearchChatMessages_MatchesBodyAndSenderCaseInsensitively()
    {
        _graphClient.GetAsync(Arg.Any<string>(), Arg.Any<Dictionary<string, string>?>())
            .Returns(Page(null, ("Q3 Budget review", "Ann"), ("lunch?", "Budget Bot"), ("unrelated", "Bob")));

        var result = await _tools.SearchChatMessages("chat1", "budget");

        Assert.Equal(2, ResultCount(result));
    }

    [Fact]
    public async Task SearchChannelMessages_FollowsNextLinkUntilEnoughMatches()
    {
        const string next = "https://graph.microsoft.com/v1.0/teams/t/channels/c/messages?$skiptoken=abc";
        _graphClient.GetAsync("teams/t/channels/c/messages", Arg.Any<Dictionary<string, string>?>())
            .Returns(Page(next, ("budget one", "Ann")));
        _graphClient.GetAsync(next, Arg.Any<Dictionary<string, string>?>())
            .Returns(Page(null, ("budget two", "Ann"), ("budget three", "Ann")));

        var result = await _tools.SearchChannelMessages("t", "c", "budget", count: 2);

        Assert.Equal(2, ResultCount(result));
        await _graphClient.Received(1).GetAsync(next, null);
    }

    [Fact]
    public async Task SearchUsers_RemovesDoubleQuotesFromQuery()
    {
        _graphClient.GetAsync(Arg.Any<string>(), Arg.Any<Dictionary<string, string>?>())
            .Returns("{\"value\":[]}");

        await _tools.SearchUsers("John \"JJ\" Smith");

        await _graphClient.Received(1).GetAsync(
            "users",
            Arg.Is<Dictionary<string, string>>(d =>
                d["$search"] == "\"displayName:John JJ Smith\" OR \"mail:John JJ Smith\" OR \"userPrincipalName:John JJ Smith\""));
    }
}
