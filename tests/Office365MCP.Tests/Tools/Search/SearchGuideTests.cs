using Office365MCP.Tools.Search;
using Xunit;

namespace Office365MCP.Tests.Tools.Search;

public class SearchGuideTests
{
    private static readonly SearchFilters NoFilters = new(From: false, Subject: false, Dates: false, SupportsSubject: true);

    private static List<SearchHit> Hits(params (string address, string date)[] hits) =>
        hits.Select(h => new SearchHit(null, h.address, DateTimeOffset.Parse(h.date))).ToList();

    [Fact]
    public void Hint_OneSenderDominatesAndMoreExist_SuggestsFromFilter()
    {
        var hits = Hits(("a@x.com", "2026-10-01"), ("a@x.com", "2026-10-02"), ("a@x.com", "2026-09-03"), ("b@x.com", "2026-09-04"));

        var hint = SearchGuide.Hint(hits, total: 312, moreAvailable: true, NoFilters);

        Assert.Contains("312 matches, showing 1 to 4.", hint);
        Assert.Contains("from=\"a@x.com\"", hint);
        Assert.Contains("2026-09-03 to 2026-10-02", hint);
    }

    [Fact]
    public void Hint_FromAlreadySet_DoesNotSuggestIt()
    {
        var hits = Hits(("a@x.com", "2026-10-01"), ("a@x.com", "2026-10-02"), ("a@x.com", "2026-10-03"), ("a@x.com", "2026-10-04"));

        var hint = SearchGuide.Hint(hits, total: null, moreAvailable: true, NoFilters with { From = true });

        Assert.DoesNotContain("from=", hint);
    }

    [Fact]
    public void Hint_FollowUpPage_OnlyOffersNextPage()
    {
        var hits = Hits(("a@x.com", "2026-10-01"));

        var hint = SearchGuide.Hint(hits, total: null, moreAvailable: true, filters: null);

        Assert.Equal("More matches exist than the 1 shown. Or pass nextPage to continue.", hint);
    }

    [Fact]
    public void Hint_NothingMore_SaysAllShown()
    {
        var hint = SearchGuide.Hint(Hits(("a@x.com", "2026-10-01")), total: 1, moreAvailable: false, NoFilters);

        Assert.Equal("All matches are shown.", hint);
    }

    [Fact]
    public void Hint_NoMatchesWithFilters_SuggestsLoosening()
    {
        var hint = SearchGuide.Hint(new List<SearchHit>(), total: 0, moreAvailable: false, NoFilters with { Dates = true });

        Assert.Contains("remove a filter", hint);
    }
}
