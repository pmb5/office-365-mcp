using System.Globalization;

namespace Office365MCP.Tools.Search;

public static class KqlQuery
{
    // dateProperty differs by source: "received" for mail, "sent" for Teams messages.
    public static string Build(
        string? text,
        string? from,
        string? subject,
        string? startDate,
        string? endDate,
        string dateProperty,
        bool? hasAttachments = null)
    {
        var clauses = new List<string>();

        if (!string.IsNullOrWhiteSpace(text))
            clauses.Add(text.Trim());
        if (!string.IsNullOrWhiteSpace(from))
            clauses.Add($"from:{QuoteIfNeeded(from)}");
        if (!string.IsNullOrWhiteSpace(subject))
            clauses.Add($"subject:\"{StripQuotes(subject)}\"");
        if (!string.IsNullOrWhiteSpace(startDate))
            clauses.Add($"{dateProperty}>={ParseDate(startDate, nameof(startDate))}");
        if (!string.IsNullOrWhiteSpace(endDate))
            clauses.Add($"{dateProperty}<={ParseDate(endDate, nameof(endDate))}");
        if (hasAttachments.HasValue)
            clauses.Add($"hasattachments:{(hasAttachments.Value ? "true" : "false")}");

        if (clauses.Count == 0)
            throw new ArgumentException("Give a query or at least one filter.");

        return string.Join(" AND ", clauses);
    }

    private static string ParseDate(string value, string name)
    {
        if (!DateOnly.TryParseExact(value.Trim(), "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
            throw new ArgumentException($"{name} must be a date like 2026-09-01, got '{value}'.");
        return date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
    }

    private static string QuoteIfNeeded(string value)
    {
        var stripped = StripQuotes(value);
        return stripped.Any(char.IsWhiteSpace) ? $"\"{stripped}\"" : stripped;
    }

    // A stray double quote would end the quoted value early and change the meaning of the query.
    private static string StripQuotes(string value) => value.Replace("\"", "").Trim();
}
