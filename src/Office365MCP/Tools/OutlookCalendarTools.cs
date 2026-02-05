using System.ComponentModel;
using System.Text.Json;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Server;
using Office365MCP.Graph;

namespace Office365MCP.Tools;

[McpServerToolType]
public class OutlookCalendarTools
{
    private readonly IGraphClientService _graphClient;
    private readonly ILogger<OutlookCalendarTools> _logger;

    public OutlookCalendarTools(IGraphClientService graphClient, ILogger<OutlookCalendarTools> logger)
    {
        _graphClient = graphClient;
        _logger = logger;
    }

    [McpServerTool, Description("List calendar events within a specified date range")]
    public async Task<string> ListCalendarEvents(
        [Description("Start date/time in ISO 8601 format (e.g., '2024-01-01T00:00:00Z')")] string startDateTime,
        [Description("End date/time in ISO 8601 format (e.g., '2024-01-31T23:59:59Z')")] string endDateTime)
    {
        try
        {
            var queryParams = new Dictionary<string, string>
            {
                ["startDateTime"] = startDateTime,
                ["endDateTime"] = endDateTime,
                ["$select"] = "id,subject,start,end,location,attendees,organizer,isOnlineMeeting,onlineMeetingUrl",
                ["$orderby"] = "start/dateTime"
            };

            var response = await _graphClient.GetAsync("/me/calendarView", queryParams);
            return response;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to list calendar events");
            return JsonSerializer.Serialize(new { error = ex.Message });
        }
    }

    [McpServerTool, Description("Get a specific calendar event by its ID")]
    public async Task<string> GetCalendarEvent(
        [Description("The ID of the calendar event")] string eventId)
    {
        try
        {
            var response = await _graphClient.GetAsync($"/me/events/{eventId}");
            return response;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to get calendar event {EventId}", eventId);
            return JsonSerializer.Serialize(new { error = ex.Message });
        }
    }

    [McpServerTool, Description("Create a new calendar event")]
    public async Task<string> CreateCalendarEvent(
        [Description("Event subject/title")] string subject,
        [Description("Start date/time in ISO 8601 format (e.g., '2024-01-01T09:00:00')")] string startDateTime,
        [Description("End date/time in ISO 8601 format (e.g., '2024-01-01T10:00:00')")] string endDateTime,
        [Description("Time zone (e.g., 'Pacific Standard Time', default: UTC)")] string timeZone = "UTC",
        [Description("Event location (optional)")] string? location = null,
        [Description("Comma-separated list of attendee email addresses (optional)")] string? attendees = null,
        [Description("Event body/description (optional)")] string? body = null)
    {
        try
        {
            var eventObject = new Dictionary<string, object>
            {
                ["subject"] = subject,
                ["start"] = new { dateTime = startDateTime, timeZone },
                ["end"] = new { dateTime = endDateTime, timeZone }
            };

            if (!string.IsNullOrEmpty(location))
            {
                eventObject["location"] = new { displayName = location };
            }

            if (!string.IsNullOrEmpty(attendees))
            {
                var attendeeList = attendees.Split(',')
                    .Select(email => new
                    {
                        emailAddress = new { address = email.Trim() },
                        type = "required"
                    })
                    .ToArray();
                eventObject["attendees"] = attendeeList;
            }

            if (!string.IsNullOrEmpty(body))
            {
                eventObject["body"] = new { contentType = "Text", content = body };
            }

            var jsonBody = JsonSerializer.Serialize(eventObject);
            var response = await _graphClient.PostAsync("/me/events", jsonBody);
            return response;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to create calendar event");
            return JsonSerializer.Serialize(new { error = ex.Message });
        }
    }

    [McpServerTool, Description("List all calendars in the user's mailbox")]
    public async Task<string> ListCalendars()
    {
        try
        {
            var response = await _graphClient.GetAsync("/me/calendars");
            return response;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to list calendars");
            return JsonSerializer.Serialize(new { error = ex.Message });
        }
    }
}
