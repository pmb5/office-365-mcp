using Microsoft.Extensions.Logging;
using NSubstitute;
using Office365MCP.Graph;
using Office365MCP.Tools;
using Xunit;

namespace Office365MCP.Tests.Tools;

public class OneNoteToolsTests
{
    private readonly IGraphClientService _graphClient;
    private readonly OneNoteTools _tools;

    public OneNoteToolsTests()
    {
        _graphClient = Substitute.For<IGraphClientService>();
        _tools = new OneNoteTools(_graphClient, Substitute.For<ILogger<OneNoteTools>>());
    }

    [Fact]
    public async Task CreatePage_PostsHtmlNotJson()
    {
        await _tools.CreatePage("s1", "Notes", "<p>hi</p>");

        await _graphClient.Received(1).PostHtmlAsync("me/onenote/sections/s1/pages", Arg.Is<string>(h => h.Contains("<p>hi</p>")));
        await _graphClient.DidNotReceive().PostAsync(Arg.Any<string>(), Arg.Any<string>());
    }

    [Fact]
    public async Task CreatePage_EncodesTitle()
    {
        await _tools.CreatePage("s1", "Q&A <draft>", "<p>hi</p>");

        await _graphClient.Received(1).PostHtmlAsync(
            Arg.Any<string>(),
            Arg.Is<string>(h => h.Contains("<title>Q&amp;A &lt;draft&gt;</title>")));
    }
}
