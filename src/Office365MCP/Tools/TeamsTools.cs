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

    [McpServerTool, Description("Search messages in a specific channel")]
    public async Task<string> SearchChannelMessages(
        [Description("The ID of the team")] string teamId,
        [Description("The ID of the channel")] string channelId,
        [Description("Search query")] string query,
        [Description("Number of results to return (default: 20)")] int count = 20)
    {
        try
        {
            var queryParams = new Dictionary<string, string>
            {
                ["$search"] = $"\"{query}\"",
                ["$top"] = count.ToString()
            };

            var response = await _graphClient.GetAsync($"teams/{teamId}/channels/{channelId}/messages", queryParams);
            return response;
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

    [McpServerTool, Description("List the user's chats")]
    public async Task<string> ListChats()
    {
        try
        {
            var response = await _graphClient.GetAsync("me/chats");
            return response;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to list chats");
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

    [McpServerTool, Description("Search messages in a specific chat")]
    public async Task<string> SearchChatMessages(
        [Description("The ID of the chat")] string chatId,
        [Description("Search query")] string query,
        [Description("Number of results to return (default: 20)")] int count = 20)
    {
        try
        {
            var queryParams = new Dictionary<string, string>
            {
                ["$search"] = $"\"{query}\"",
                ["$top"] = count.ToString()
            };

            var response = await _graphClient.GetAsync($"chats/{chatId}/messages", queryParams);
            return response;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to search messages in chat {ChatId}", chatId);
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
