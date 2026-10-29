using System.Globalization;

namespace Office365MCP.Tools.Search;

// Builds the summary and next-step hint returned with every search page, so the agent can
// narrow a broad search instead of fetching more pages. Works only on the page already
// downloaded: it never makes extra Graph requests.
public static class SearchGuide
{
    private const int TopSenderCount = 3;

    public static object Summarize(IReadOnlyList<SearchHit> hits)
    {
        var dates = hits.Where(h => h.Date.HasValue).Select(h => h.Date!.Value).ToList();

        return new
        {
            topSenders = TopSenders(hits)
                .Select(g => new { name = g.Name, address = g.Address, count = g.Count }),
            earliest = dates.Count > 0 ? dates.Min().ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) : null,
            latest = dates.Count > 0 ? dates.Max().ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) : null,
            byMonth = dates
                .GroupBy(d => d.ToString("yyyy-MM", CultureInfo.InvariantCulture))
                .OrderByDescending(g => g.Key)
                .Select(g => new { month = g.Key, count = g.Count() })
        };
    }

    // filters is null for a follow-up page, where the original filters are not known.
    public static string Hint(IReadOnlyList<SearchHit> hits, long? total, bool moreAvailable, SearchFilters? filters)
    {
        if (hits.Count == 0)
        {
            return filters is { From: false, Subject: false, Dates: false }
                ? "No matches. Try fewer or different words."
                : "No matches. Try fewer words, remove a filter, or widen the date range.";
        }

        if (!moreAvailable)
            return "All matches are shown.";

        var parts = new List<string>
        {
            total.HasValue
                ? $"{total} matches, showing {hits.Count}."
                : $"More matches exist than the {hits.Count} shown."
        };

        if (filters != null)
        {
            var top = TopSenders(hits).FirstOrDefault();
            if (!filters.From && top != null && hits.Count >= 4 && top.Count * 2 >= hits.Count)
                parts.Add($"{top.Count} of these {hits.Count} are from {top.Name ?? top.Address}; to focus on them set from=\"{top.Address ?? top.Name}\".");

            if (!filters.Dates)
            {
                var dates = hits.Where(h => h.Date.HasValue).Select(h => h.Date!.Value).ToList();
                parts.Add(dates.Count > 0
                    ? $"Set startDate/endDate to narrow by date; this page spans {dates.Min():yyyy-MM-dd} to {dates.Max():yyyy-MM-dd}."
                    : "Set startDate/endDate to narrow by date.");
            }

            if (filters.SupportsSubject && !filters.Subject)
                parts.Add("Set subject if you know words in the subject.");
        }

        parts.Add("Or pass nextPage to continue.");
        return string.Join(" ", parts);
    }

    private static List<SenderCount> TopSenders(IReadOnlyList<SearchHit> hits) =>
        hits
            .Where(h => h.SenderAddress != null || h.SenderName != null)
            .GroupBy(h => (h.SenderAddress ?? h.SenderName)!.ToLowerInvariant())
            .Select(g => new SenderCount(g.First().SenderName, g.First().SenderAddress, g.Count()))
            .OrderByDescending(s => s.Count)
            .Take(TopSenderCount)
            .ToList();

    private record SenderCount(string? Name, string? Address, int Count);
}
