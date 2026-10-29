using System.ComponentModel;
using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Server;
using Office365MCP.Graph;
using Office365MCP.Tools.Search;

namespace Office365MCP.Tools;

[McpServerToolType]
public class OutlookMailTools
{
    private readonly IGraphClientService _graphClient;
    private readonly ILogger<OutlookMailTools> _logger;

    public OutlookMailTools(IGraphClientService graphClient, ILogger<OutlookMailTools> logger)
    {
        _graphClient = graphClient;
        _logger = logger;
    }

    [McpServerTool, Description("List recent emails from the user's mailbox")]
    public async Task<string> ListEmails(
        [Description("Number of emails to retrieve (default: 10)")] int count = 10)
    {
        try
        {
            var queryParams = new Dictionary<string, string>
            {
                ["$top"] = count.ToString(),
                ["$select"] = "id,subject,from,receivedDateTime,bodyPreview,isRead",
                ["$orderby"] = "receivedDateTime desc"
            };

            var response = await _graphClient.GetAsync("me/messages", queryParams);
            return response;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to list emails");
            return JsonSerializer.Serialize(new { error = ex.Message });
        }
    }

    [McpServerTool, Description("Get a specific email by its ID")]
    public async Task<string> GetEmail(
        [Description("The ID of the email message")] string messageId)
    {
        try
        {
            var response = await _graphClient.GetAsync($"me/messages/{messageId}");
            return response;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to get email {MessageId}", messageId);
            return JsonSerializer.Serialize(new { error = ex.Message });
        }
    }

    [McpServerTool, Description(
        "Search the mailbox, newest first. Start narrow: prefer the from, subject, startDate, endDate and " +
        "hasAttachments filters over writing search syntax. Each response reports who the results are from, " +
        "which dates they span, and a hint on how to narrow the search or get the next page. Follow the hint " +
        "rather than raising count.")]
    public async Task<string> SearchEmails(
        [Description("Words to find anywhere in the email. Optional when a filter is set. Also accepts KQL, e.g. body:invoice or to:jane@example.com.")] string? query = null,
        [Description("Sender name or email address")] string? from = null,
        [Description("Words or a phrase in the subject")] string? subject = null,
        [Description("Only emails received on or after this date (yyyy-MM-dd)")] string? startDate = null,
        [Description("Only emails received on or before this date (yyyy-MM-dd)")] string? endDate = null,
        [Description("true for only emails with attachments, false for only emails without")] bool? hasAttachments = null,
        [Description("Number of results per page (default: 10)")] int count = 10,
        [Description("The nextPage token from a previous response, to get the next page of that search. Other parameters are ignored when it is set.")] string? nextPage = null)
    {
        try
        {
            string response;
            string? kql = null;
            SearchFilters? filters = null;

            if (nextPage != null)
            {
                var link = SearchCursor.Decode(nextPage);
                // The Graph client attaches the user's token, so only follow links back to the same search.
                if (!link.StartsWith(MessagesNextLinkPrefix, StringComparison.OrdinalIgnoreCase))
                    throw new ArgumentException("nextPage is not a token from a previous email search.");
                response = await _graphClient.GetAsync(link);
            }
            else
            {
                kql = KqlQuery.Build(query, from, subject, startDate, endDate, "received", hasAttachments);
                filters = new SearchFilters(
                    From: from != null,
                    Subject: subject != null,
                    Dates: startDate != null || endDate != null,
                    SupportsSubject: true);

                // Graph rejects $search combined with $orderby on messages (400 Bad Request);
                // $search results already come back newest first. Inner double quotes must be
                // backslash-escaped, otherwise a quoted phrase in the query breaks the request.
                var escapedQuery = kql.Replace("\\", "\\\\").Replace("\"", "\\\"");
                var queryParams = new Dictionary<string, string>
                {
                    ["$search"] = $"\"{escapedQuery}\"",
                    ["$top"] = count.ToString(),
                    ["$select"] = "id,subject,from,receivedDateTime,bodyPreview,isRead,hasAttachments"
                };

                response = await _graphClient.GetAsync("me/messages", queryParams);
            }

            var data = JsonSerializer.Deserialize<JsonElement>(response);
            var results = new List<object>();
            var hits = new List<SearchHit>();

            if (data.TryGetProperty("value", out var messages))
            {
                foreach (var message in messages.EnumerateArray())
                {
                    var name = JsonPath.GetString(message, "from", "emailAddress", "name");
                    var address = JsonPath.GetString(message, "from", "emailAddress", "address");
                    var received = DateTimeOffset.TryParse(JsonPath.GetString(message, "receivedDateTime"), out var r) ? r : (DateTimeOffset?)null;

                    hits.Add(new SearchHit(name, address, received));
                    results.Add(new
                    {
                        id = JsonPath.GetString(message, "id"),
                        subject = JsonPath.GetString(message, "subject"),
                        from = new { name, address },
                        received = JsonPath.GetString(message, "receivedDateTime"),
                        isRead = message.TryGetProperty("isRead", out var isRead) && isRead.ValueKind == JsonValueKind.True,
                        hasAttachments = message.TryGetProperty("hasAttachments", out var att) && att.ValueKind == JsonValueKind.True,
                        preview = CleanPreview(JsonPath.GetString(message, "bodyPreview"))
                    });
                }
            }

            var nextLink = data.TryGetProperty("@odata.nextLink", out var next) ? next.GetString() : null;

            return JsonSerializer.Serialize(new
            {
                query = kql,
                returned = hits.Count,
                moreAvailable = nextLink != null,
                nextPage = nextLink != null ? SearchCursor.Encode(nextLink) : null,
                summary = SearchGuide.Summarize(hits),
                hint = SearchGuide.Hint(hits, total: null, moreAvailable: nextLink != null, filters),
                results
            }, SearchJson.Options);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to search emails with query: {Query}", query);
            return JsonSerializer.Serialize(new { error = ex.Message });
        }
    }

    private const string MessagesNextLinkPrefix = "https://graph.microsoft.com/v1.0/me/messages?";
    private const int PreviewLength = 200;

    // Marketing mail pads bodyPreview with invisible spacing characters that cost tokens and say nothing.
    private static string? CleanPreview(string? preview)
    {
        if (preview == null)
            return null;

        var visible = new string(preview
            .Where(c => !char.IsControl(c) || char.IsWhiteSpace(c))
            .Where(c => CharUnicodeInfo.GetUnicodeCategory(c) is not (UnicodeCategory.NonSpacingMark or UnicodeCategory.Format))
            .Select(c => char.IsWhiteSpace(c) ? ' ' : c)
            .ToArray());
        var collapsed = Regex.Replace(visible, " {2,}", " ").Trim();
        return collapsed.Length <= PreviewLength ? collapsed : collapsed[..PreviewLength] + "…";
    }

    [McpServerTool, Description("Send an email")]
    public async Task<string> SendEmail(
        [Description("Recipient email address")] string to,
        [Description("Email subject")] string subject,
        [Description("Email body (plain text or HTML)")] string body,
        [Description("Content type: 'Text' or 'HTML' (default: Text)")] string contentType = "Text")
    {
        try
        {
            var message = new
            {
                message = new
                {
                    subject,
                    body = new
                    {
                        contentType,
                        content = body
                    },
                    toRecipients = new[]
                    {
                        new
                        {
                            emailAddress = new
                            {
                                address = to
                            }
                        }
                    }
                }
            };

            var jsonBody = JsonSerializer.Serialize(message);
            var response = await _graphClient.PostAsync("me/sendMail", jsonBody);
            return JsonSerializer.Serialize(new { success = true, message = "Email sent successfully" });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to send email to {Recipient}", to);
            return JsonSerializer.Serialize(new { error = ex.Message });
        }
    }

    [McpServerTool, Description("List mail folders in the user's mailbox")]
    public async Task<string> ListMailFolders()
    {
        try
        {
            var response = await _graphClient.GetAsync("me/mailFolders");
            return response;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to list mail folders");
            return JsonSerializer.Serialize(new { error = ex.Message });
        }
    }
}
