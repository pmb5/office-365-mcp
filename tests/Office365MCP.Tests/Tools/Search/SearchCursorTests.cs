using Office365MCP.Tools.Search;
using Xunit;

namespace Office365MCP.Tests.Tools.Search;

public class SearchCursorTests
{
    [Fact]
    public void Decode_EncodedValue_RoundTrips()
    {
        const string link = "https://graph.microsoft.com/v1.0/me/messages?$search=%22a%22&$skiptoken=MSZZ+/=";

        Assert.Equal(link, SearchCursor.Decode(SearchCursor.Encode(link)));
    }

    [Fact]
    public void Decode_NotAToken_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>(() => SearchCursor.Decode("not a token!"));
    }
}
