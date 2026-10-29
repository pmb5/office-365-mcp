using Microsoft.Extensions.Logging;
using NSubstitute;
using Office365MCP.Graph;
using Office365MCP.Tools;
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
}
