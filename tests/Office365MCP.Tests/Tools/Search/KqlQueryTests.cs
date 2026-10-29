using Office365MCP.Tools.Search;
using Xunit;

namespace Office365MCP.Tests.Tools.Search;

public class KqlQueryTests
{
    [Fact]
    public void Build_WithTextAndFilters_JoinsWithAnd()
    {
        var kql = KqlQuery.Build("invoice", "billing@example.com", "Your receipt", "2026-09-01", "2026-09-30", "received", hasAttachments: true);

        Assert.Equal(
            "invoice AND from:billing@example.com AND subject:\"Your receipt\" AND received>=2026-09-01 AND received<=2026-09-30 AND hasattachments:true",
            kql);
    }

    [Fact]
    public void Build_SenderNameWithSpace_IsQuoted()
    {
        var kql = KqlQuery.Build(null, "Matthew Skibba", null, null, null, "sent");

        Assert.Equal("from:\"Matthew Skibba\"", kql);
    }

    [Fact]
    public void Build_FilterValueWithQuotes_HasQuotesRemoved()
    {
        var kql = KqlQuery.Build(null, null, "the \"big\" launch", null, null, "received");

        Assert.Equal("subject:\"the big launch\"", kql);
    }

    [Theory]
    [InlineData("09/01/2026")]
    [InlineData("last week")]
    public void Build_DateNotIsoFormatted_ThrowsWithExample(string date)
    {
        var ex = Assert.Throws<ArgumentException>(() => KqlQuery.Build("x", null, null, date, null, "received"));

        Assert.Contains("2026-09-01", ex.Message);
    }

    [Fact]
    public void Build_NoTextOrFilters_Throws()
    {
        Assert.Throws<ArgumentException>(() => KqlQuery.Build(" ", null, null, null, null, "received"));
    }
}
