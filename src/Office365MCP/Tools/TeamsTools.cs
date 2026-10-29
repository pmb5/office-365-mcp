using System.ComponentModel;
using System.Text.Json;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Server;
using Office365MCP.Graph;
using Office365MCP.Tools.Search;

namespace Office365MCP.Tools;

[McpServerToolType]
public class TeamsTools
{
    private readonly IGraphClientService _graphClient;
    private readonly ILogger<TeamsTools> _logger;

    public TeamsTools(IGraphClientService graphClient, ILogger<TeamsTools> logger)
    {
        _graphClient = graphClient;
        _logger = logger;
    }

    [McpServerTool, Description("List all teams the user has joined")]
    public async Task<string> ListJoinedTeams()
    {
        try
        {
            var response = await _graphClient.GetAsync("me/joinedTeams");
            return response;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to list joined teams");
            return JsonSerializer.Serialize(new { error = ex.Message });
        }
    }

    [McpServerTool, Description("List channels in a specific team")]
    public async Task<string> ListChannels(
        [Description("The ID of the team")] string teamId)
    {
        try
        {
            var response = await _graphClient.GetAsync($"teams/{teamId}/channels");
            return response;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to list channels for team {TeamId}", teamId);
            return JsonSerializer.Serialize(new { error = ex.Message });
        }
    }

    [McpServerTool, Description("List recent messages in a specific channel")]
    public async Task<string> ListChannelMessages(
        [Description("The ID of the team")] string teamId,
        [Description("The ID of the channel")] string channelId,
        [Description("Number of messages to retrieve (default: 20)")] int count = 20)
    {
        try
        {
            var queryParams = new Dictionary<string, string>
            {
                ["$top"] = count.ToString()
            };

            var response = await _graphClient.GetAsync($"teams/{teamId}/channels/{channelId}/messages", queryParams);
            return response;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to list messages for channel {ChannelId} in team {TeamId}", channelId, teamId);
            return JsonSerializer.Serialize(new { error = ex.Message });
        }
    }

    [McpServerTool, Description(
        "Search messages across all Teams chats and channels, including older history. Start narrow: prefer " +
        "the from, startDate and endDate filters over writing search syntax. Each result names the chat or " +
        "channel it came from. Each response reports the total number of matches, who they are from, which " +
        "dates they span, and a hint on how to narrow the search or get the next page. Follow the hint rather " +
        "than raising count.")]
    public async Task<string> SearchTeamsMessages(
        [Description("Words to find in the message. Optional when a filter is set.")] string? query = null,
        [Description("Sender name or email address")] string? from = null,
        [Description("Only messages sent on or after this date (yyyy-MM-dd)")] string? startDate = null,
        [Description("Only messages sent on or before this date (yyyy-MM-dd)")] string? endDate = null,
        [Description("Number of results per page (default: 20)")] int count = 20,
        [Description("The nextPage token from a previous response, to get the next page of that search. Other parameters are ignored when it is set.")] string? nextPage = null)
    {
        try
        {
            string kql;
            int offset;
            SearchFilters? filters = null;

            if (nextPage != null)
            {
                // The Search API pages by offset, so the token carries the query and where to resume.
                var page = JsonSerializer.Deserialize<TeamsSearchPage>(SearchCursor.Decode(nextPage));
                if (page == null || string.IsNullOrEmpty(page.Query))
                    throw new ArgumentException("nextPage is not a token from a previous Teams search.");
                kql = page.Query;
                offset = page.Offset;
            }
            else
            {
                kql = KqlQuery.Build(query, from, subject: null, startDate, endDate, "sent");
                offset = 0;
                filters = new SearchFilters(
                    From: from != null,
                    Subject: false,
                    Dates: startDate != null || endDate != null,
                    SupportsSubject: false);
            }

            var searchRequest = new
            {
                requests = new[]
                {
                    new
                    {
                        entityTypes = new[] { "chatMessage" },
                        query = new
                        {
                            queryString = kql
                        },
                        from = offset,
                        size = count
                    }
                }
            };

            var jsonBody = JsonSerializer.Serialize(searchRequest);
            var response = await _graphClient.PostAsync("search/query", jsonBody);

            var data = JsonSerializer.Deserialize<JsonElement>(response);
            var results = new List<object>();
            var hits = new List<SearchHit>();
            long? total = null;
            var moreAvailable = false;

            if (data.TryGetProperty("value", out var responses) && responses.GetArrayLength() > 0
                && responses[0].TryGetProperty("hitsContainers", out var containers) && containers.GetArrayLength() > 0)
            {
                var container = containers[0];
                total = container.TryGetProperty("total", out var t) && t.ValueKind == JsonValueKind.Number ? t.GetInt64() : null;
                moreAvailable = container.TryGetProperty("moreResultsAvailable", out var more) && more.ValueKind == JsonValueKind.True;

                if (container.TryGetProperty("hits", out var hitList))
                {
                    foreach (var hit in hitList.EnumerateArray())
                    {
                        if (!hit.TryGetProperty("resource", out var resource))
                            continue;

                        var name = JsonPath.GetString(resource, "from", "emailAddress", "name");
                        var address = JsonPath.GetString(resource, "from", "emailAddress", "address");
                        var sent = JsonPath.GetString(resource, "createdDateTime");
                        var teamId = JsonPath.GetString(resource, "channelIdentity", "teamId");

                        hits.Add(new SearchHit(name, address, DateTimeOffset.TryParse(sent, out var s) ? s : null));
                        results.Add(new
                        {
                            id = JsonPath.GetString(resource, "id"),
                            chatId = JsonPath.GetString(resource, "chatId"),
                            // Chat hits repeat the chat ID as a channel ID; only real channels have a team.
                            teamId,
                            channelId = teamId != null ? JsonPath.GetString(resource, "channelIdentity", "channelId") : null,
                            from = new { name, address },
                            sent,
                            summary = JsonPath.GetString(hit, "summary"),
                            webLink = JsonPath.GetString(resource, "webLink")
                        });
                    }
                }
            }

            return JsonSerializer.Serialize(new
            {
                query = kql,
                returned = hits.Count,
                total,
                moreAvailable,
                nextPage = moreAvailable
                    ? SearchCursor.Encode(JsonSerializer.Serialize(new TeamsSearchPage(kql, offset + hits.Count)))
                    : null,
                summary = SearchGuide.Summarize(hits),
                hint = SearchGuide.Hint(hits, total, moreAvailable, filters),
                results
            }, SearchJson.Options);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to search Teams messages with query: {Query}", query);
            return JsonSerializer.Serialize(new { error = ex.Message });
        }
    }

    [McpServerTool, Description("Send a message to a channel")]
    public async Task<string> SendChannelMessage(
        [Description("The ID of the team")] string teamId,
        [Description("The ID of the channel")] string channelId,
        [Description("Message content (supports HTML)")] string message)
    {
        try
        {
            var messageObject = new
            {
                body = new
                {
                    contentType = "html",
                    content = message
                }
            };

            var jsonBody = JsonSerializer.Serialize(messageObject);
            var response = await _graphClient.PostAsync($"teams/{teamId}/channels/{channelId}/messages", jsonBody);
            return response;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to send message to channel {ChannelId}", channelId);
            return JsonSerializer.Serialize(new { error = ex.Message });
        }
    }

    [McpServerTool, Description("List the user's chats with member information")]
    public async Task<string> ListChats()
    {
        try
        {
            // Include members to see who's in each chat
            var queryParams = new Dictionary<string, string>
            {
                ["$expand"] = "members"
            };
            var response = await _graphClient.GetAsync("me/chats", queryParams);
            return response;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to list chats");
            return JsonSerializer.Serialize(new { error = ex.Message });
        }
    }

    [McpServerTool, Description("Search for users in the organization by name or email")]
    public async Task<string> SearchUsers(
        [Description("Search query - can be a full name, partial name, or email address (e.g., 'John Smith' or 'john@')")] string query,
        [Description("Number of results to return (default: 10)")] int count = 10)
    {
        try
        {
            // A double quote would end the quoted search clause early and make Graph reject the request.
            var term = query.Replace("\"", "");
            var queryParams = new Dictionary<string, string>
            {
                ["$search"] = $"\"displayName:{term}\" OR \"mail:{term}\" OR \"userPrincipalName:{term}\"",
                ["$top"] = count.ToString(),
                ["$select"] = "id,displayName,mail,userPrincipalName,jobTitle,department"
            };

            var response = await _graphClient.GetAsync("users", queryParams);
            return response;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to search users with query: {Query}", query);
            return JsonSerializer.Serialize(new { error = ex.Message });
        }
    }

    [McpServerTool, Description("Find the one-on-one chat ID with a specific user")]
    public async Task<string> FindChatWithUser(
        [Description("The email address or user ID of the person you want to chat with")] string userEmailOrId)
    {
        try
        {
            // Get all chats with members expanded
            var queryParams = new Dictionary<string, string>
            {
                ["$expand"] = "members",
                ["$filter"] = "chatType eq 'oneOnOne'"
            };

            var response = await _graphClient.GetAsync("me/chats", queryParams);
            var chatsData = JsonSerializer.Deserialize<JsonElement>(response);

            if (!chatsData.TryGetProperty("value", out var chats))
            {
                return JsonSerializer.Serialize(new { error = "No chats found" });
            }

            // Look through each chat to find the one with the specified user
            foreach (var chat in chats.EnumerateArray())
            {
                if (!chat.TryGetProperty("members", out var members))
                    continue;

                foreach (var member in members.EnumerateArray())
                {
                    if (!member.TryGetProperty("userId", out var userId))
                        continue;

                    var userIdStr = userId.GetString() ?? "";

                    // Also check email if available
                    var email = "";
                    if (member.TryGetProperty("email", out var emailProp))
                    {
                        email = emailProp.GetString() ?? "";
                    }

                    // Match by user ID or email
                    if (userIdStr.Equals(userEmailOrId, StringComparison.OrdinalIgnoreCase) ||
                        email.Equals(userEmailOrId, StringComparison.OrdinalIgnoreCase))
                    {
                        if (chat.TryGetProperty("id", out var chatId))
                        {
                            var displayName = "";
                            if (member.TryGetProperty("displayName", out var displayNameProp))
                            {
                                displayName = displayNameProp.GetString() ?? "";
                            }

                            return JsonSerializer.Serialize(new
                            {
                                chatId = chatId.GetString(),
                                displayName,
                                email,
                                userId = userIdStr,
                                chatType = chat.GetProperty("chatType").GetString()
                            });
                        }
                    }
                }
            }

            return JsonSerializer.Serialize(new
            {
                error = $"No one-on-one chat found with user: {userEmailOrId}",
                suggestion = "The user might not have an existing chat. You may need to start a new conversation."
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to find chat with user {UserEmailOrId}", userEmailOrId);
            return JsonSerializer.Serialize(new { error = ex.Message });
        }
    }

    [McpServerTool, Description("List messages in a specific chat")]
    public async Task<string> ListChatMessages(
        [Description("The ID of the chat")] string chatId,
        [Description("Number of messages to retrieve (default: 20)")] int count = 20)
    {
        try
        {
            var queryParams = new Dictionary<string, string>
            {
                ["$top"] = count.ToString()
            };

            var response = await _graphClient.GetAsync($"chats/{chatId}/messages", queryParams);
            return response;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to list messages for chat {ChatId}", chatId);
            return JsonSerializer.Serialize(new { error = ex.Message });
        }
    }

    [McpServerTool, Description("Send a message to a chat")]
    public async Task<string> SendChatMessage(
        [Description("The ID of the chat")] string chatId,
        [Description("Message content (supports HTML)")] string message)
    {
        try
        {
            var messageObject = new
            {
                body = new
                {
                    contentType = "html",
                    content = message
                }
            };

            var jsonBody = JsonSerializer.Serialize(messageObject);
            var response = await _graphClient.PostAsync($"chats/{chatId}/messages", jsonBody);
            return response;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to send message to chat {ChatId}", chatId);
            return JsonSerializer.Serialize(new { error = ex.Message });
        }
    }
}
