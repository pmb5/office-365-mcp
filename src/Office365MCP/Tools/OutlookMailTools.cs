using System.ComponentModel;
using System.Text.Json;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Server;
using Office365MCP.Graph;

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

            var response = await _graphClient.GetAsync("/me/messages", queryParams);
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
            var response = await _graphClient.GetAsync($"/me/messages/{messageId}");
            return response;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to get email {MessageId}", messageId);
            return JsonSerializer.Serialize(new { error = ex.Message });
        }
    }

    [McpServerTool, Description("Search emails using Keyword Query Language (KQL)")]
    public async Task<string> SearchEmails(
        [Description("Search query using KQL syntax (e.g., 'from:john@example.com' or 'subject:meeting')")] string query,
        [Description("Number of results to return (default: 10)")] int count = 10)
    {
        try
        {
            var queryParams = new Dictionary<string, string>
            {
                ["$search"] = $"\"{query}\"",
                ["$top"] = count.ToString(),
                ["$select"] = "id,subject,from,receivedDateTime,bodyPreview,isRead",
                ["$orderby"] = "receivedDateTime desc"
            };

            var response = await _graphClient.GetAsync("/me/messages", queryParams);
            return response;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to search emails with query: {Query}", query);
            return JsonSerializer.Serialize(new { error = ex.Message });
        }
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
            var response = await _graphClient.PostAsync("/me/sendMail", jsonBody);
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
            var response = await _graphClient.GetAsync("/me/mailFolders");
            return response;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to list mail folders");
            return JsonSerializer.Serialize(new { error = ex.Message });
        }
    }
}
