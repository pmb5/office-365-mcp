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

    private static string SearchResponse(long total, bool more, params (string id, string chatId, string? teamId, string sender, string sent)[] hits) =>
        JsonSerializer.Serialize(new
        {
            value = new[]
            {
                new
                {
                    hitsContainers = new[]
                    {
                        new
                        {
                            total,
                            moreResultsAvailable = more,
                            hits = hits.Select(h => new
                            {
                                summary = "text",
                                resource = new
                                {
                                    id = h.id,
                                    chatId = h.chatId,
                                    createdDateTime = h.sent,
                                    from = new { emailAddress = new { name = h.sender, address = $"{h.sender}@x.com" } },
                                    channelIdentity = h.teamId == null
                                        ? (object)new { channelId = h.chatId }
                                        : new { teamId = h.teamId, channelId = "channel1" }
                                }
                            })
                        }
                    }
                }
            }
        });

    [Fact]
    public async Task SearchTeamsMessages_WithFilters_SendsCombinedQuery()
    {
        _graphClient.PostAsync(Arg.Any<string>(), Arg.Any<string>()).Returns(SearchResponse(0, false));

        await _tools.SearchTeamsMessages("outage", from: "Matthew Skibba", startDate: "2026-10-01");

        await _graphClient.Received(1).PostAsync(
            "search/query",
            Arg.Is<string>(body => QueryString(body) == "outage AND from:\"Matthew Skibba\" AND sent>=2026-10-01"));
    }

    private static string? QueryString(string body) =>
        JsonSerializer.Deserialize<JsonElement>(body).GetProperty("requests")[0].GetProperty("query").GetProperty("queryString").GetString();

    [Fact]
    public async Task SearchTeamsMessages_ChatHit_HasNoChannelId()
    {
        _graphClient.PostAsync(Arg.Any<string>(), Arg.Any<string>())
            .Returns(SearchResponse(1, false, ("m1", "chat1", null, "ann", "2026-10-01T10:00:00Z")));

        var result = JsonSerializer.Deserialize<JsonElement>(await _tools.SearchTeamsMessages("x"));

        Assert.Equal(JsonValueKind.Null, result.GetProperty("results")[0].GetProperty("channelId").ValueKind);
    }

    [Fact]
    public async Task SearchTeamsMessages_NextPage_ResumesAfterReturnedHits()
    {
        _graphClient.PostAsync(Arg.Any<string>(), Arg.Any<string>())
            .Returns(SearchResponse(50, true,
                ("m1", "chat1", null, "ann", "2026-10-01T10:00:00Z"),
                ("m2", "chat1", null, "ann", "2026-10-02T10:00:00Z")));

        var first = JsonSerializer.Deserialize<JsonElement>(await _tools.SearchTeamsMessages("outage", count: 2));
        await _tools.SearchTeamsMessages(nextPage: first.GetProperty("nextPage").GetString());

        await _graphClient.Received(1).PostAsync(
            "search/query",
            Arg.Is<string>(body => body.Contains("\"queryString\":\"outage\"") && body.Contains("\"from\":2")));
    }
}
