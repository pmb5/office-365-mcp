using System.ComponentModel;
using System.Text.Json;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Server;
using Office365MCP.Graph;

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

    [McpServerTool, Description("Search the most recent 200 messages in a specific channel for text in the message body or sender name. For older messages use search_teams_messages.")]
    public async Task<string> SearchChannelMessages(
        [Description("The ID of the team")] string teamId,
        [Description("The ID of the channel")] string channelId,
        [Description("Text to look for (case-insensitive)")] string query,
        [Description("Number of results to return (default: 20)")] int count = 20)
    {
        try
        {
            return await FilterRecentMessagesAsync($"teams/{teamId}/channels/{channelId}/messages", query, count);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to search messages in channel {ChannelId}", channelId);
            return JsonSerializer.Serialize(new { error = ex.Message });
        }
    }

    [McpServerTool, Description("Search messages across all Teams channels and chats")]
    public async Task<string> SearchTeamsMessages(
        [Description("Search query")] string query,
        [Description("Number of results to return (default: 20)")] int count = 20)
    {
        try
        {
            var searchRequest = new
            {
                requests = new[]
                {
                    new
                    {
                        entityTypes = new[] { "chatMessage" },
                        query = new
                        {
                            queryString = query
                        },
                        from = 0,
                        size = count
                    }
                }
            };

            var jsonBody = JsonSerializer.Serialize(searchRequest);
            var response = await _graphClient.PostAsync("search/query", jsonBody);
            return response;
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

    [McpServerTool, Description("Search the most recent 200 messages in a specific chat for text in the message body or sender name. For older messages use search_teams_messages.")]
    public async Task<string> SearchChatMessages(
        [Description("The ID of the chat")] string chatId,
        [Description("Text to look for (case-insensitive)")] string query,
        [Description("Number of results to return (default: 20)")] int count = 20)
    {
        try
        {
            return await FilterRecentMessagesAsync($"chats/{chatId}/messages", query, count);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to search messages in chat {ChatId}", chatId);
            return JsonSerializer.Serialize(new { error = ex.Message });
        }
    }

    // Graph does not support $search on channel or chat message lists, so page through
    // the most recent messages and match locally.
    private async Task<string> FilterRecentMessagesAsync(string endpoint, string query, int count)
    {
        const int pageSize = 50;
        const int maxPages = 4;

        var matches = new List<JsonElement>();
        string? next = endpoint;
        var queryParams = new Dictionary<string, string> { ["$top"] = pageSize.ToString() };

        for (var page = 0; page < maxPages && next != null && matches.Count < count; page++)
        {
            // nextLink already carries its own query string.
            var response = await _graphClient.GetAsync(next, page == 0 ? queryParams : null);
            var data = JsonSerializer.Deserialize<JsonElement>(response);

            if (data.TryGetProperty("value", out var messages))
            {
                foreach (var message in messages.EnumerateArray())
                {
                    if (MessageMatches(message, query))
                    {
                        matches.Add(message.Clone());
                        if (matches.Count == count)
                            break;
                    }
                }
            }

            next = data.TryGetProperty("@odata.nextLink", out var link) ? link.GetString() : null;
        }

        return JsonSerializer.Serialize(new { value = matches });
    }

    private static bool MessageMatches(JsonElement message, string query)
    {
        var body = message.TryGetProperty("body", out var b) && b.TryGetProperty("content", out var c) && c.ValueKind == JsonValueKind.String
            ? c.GetString()
            : null;
        var sender = message.TryGetProperty("from", out var f) && f.ValueKind == JsonValueKind.Object
            && f.TryGetProperty("user", out var u) && u.ValueKind == JsonValueKind.Object
            && u.TryGetProperty("displayName", out var n) && n.ValueKind == JsonValueKind.String
            ? n.GetString()
            : null;

        return (body?.Contains(query, StringComparison.OrdinalIgnoreCase) ?? false)
            || (sender?.Contains(query, StringComparison.OrdinalIgnoreCase) ?? false);
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
