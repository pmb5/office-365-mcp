using System.Text.Json;
using Microsoft.Extensions.Logging;
using NSubstitute;
using Office365MCP.Auth;
using Office365MCP.Tools;
using Xunit;

namespace Office365MCP.Tests.Tools;

public class PermissionManagementToolsTests
{
    private readonly IPermissionHelper _mockPermissionHelper;
    private readonly IGraphAuthService _mockAuthService;
    private readonly ILogger<PermissionManagementTools> _mockLogger;
    private readonly PermissionManagementTools _tools;

    public PermissionManagementToolsTests()
    {
        _mockPermissionHelper = Substitute.For<IPermissionHelper>();
        _mockAuthService = Substitute.For<IGraphAuthService>();
        _mockLogger = Substitute.For<ILogger<PermissionManagementTools>>();
        _tools = new PermissionManagementTools(_mockPermissionHelper, _mockAuthService, _mockLogger);
    }

    [Fact]
    public async Task ListCurrentPermissions_WithOnlyDefaults_ShowsDefaultsOnly()
    {
        // Arrange
        var defaultScopes = new[] { "User.Read", "Mail.Read", "Calendars.Read", "Chat.Read", "Notes.Read" };
        _mockPermissionHelper.GetDefaultScopes().Returns(defaultScopes);
        _mockPermissionHelper.GetCurrentScopes().Returns(defaultScopes);

        // Act
        var result = await _tools.ListCurrentPermissions();

        // Assert
        Assert.Contains("Current Permissions:", result);
        Assert.Contains("Default Permissions (cannot be reverted):", result);
        Assert.Contains("User.Read", result);
        Assert.Contains("Mail.Read", result);
        Assert.Contains("No dynamic permissions added yet.", result);
        Assert.Contains("Total: 5 permission(s) active", result);
    }

    [Fact]
    public async Task ListCurrentPermissions_WithDynamicPermissions_ShowsBothCategories()
    {
        // Arrange
        var defaultScopes = new[] { "User.Read", "Mail.Read", "Calendars.Read", "Chat.Read", "Notes.Read" };
        var currentScopes = new[] { "User.Read", "Mail.Read", "Calendars.Read", "Chat.Read", "Notes.Read", "Mail.Send", "Calendars.ReadWrite" };
        _mockPermissionHelper.GetDefaultScopes().Returns(defaultScopes);
        _mockPermissionHelper.GetCurrentScopes().Returns(currentScopes);

        // Act
        var result = await _tools.ListCurrentPermissions();

        // Assert
        Assert.Contains("Default Permissions (cannot be reverted):", result);
        Assert.Contains("Dynamically Added Permissions (can be reverted):", result);
        Assert.Contains("Mail.Send", result);
        Assert.Contains("Calendars.ReadWrite", result);
        Assert.Contains("Total: 7 permission(s) active", result);
    }

    [Fact]
    public async Task ListCurrentPermissions_WhenExceptionOccurs_ReturnsErrorJson()
    {
        // Arrange
        _mockPermissionHelper.GetDefaultScopes().Returns(x => throw new Exception("Test error"));

        // Act
        var result = await _tools.ListCurrentPermissions();

        // Assert
        Assert.Contains("error", result);
        Assert.Contains("Test error", result);
        var json = JsonSerializer.Deserialize<JsonDocument>(result);
        Assert.NotNull(json);
        Assert.True(json.RootElement.TryGetProperty("error", out _));
    }

    [Fact]
    public async Task ViewPermissionHistory_WithNoChanges_ReturnsNoChangesMessage()
    {
        // Arrange
        var emptyHistory = new PermissionChangeHistory();
        _mockPermissionHelper.GetChangeHistory().Returns(emptyHistory);

        // Act
        var result = await _tools.ViewPermissionHistory();

        // Assert
        Assert.Equal("No permission changes recorded yet.", result);
    }

    [Fact]
    public async Task ViewPermissionHistory_WithActiveAndRevertedChanges_ShowsBoth()
    {
        // Arrange
        var history = new PermissionChangeHistory
        {
            Changes = new List<PermissionChange>
            {
                new PermissionChange
                {
                    Permission = "Mail.Send",
                    Timestamp = new DateTime(2026, 2, 5, 9, 25, 10, DateTimeKind.Utc),
                    Reason = "Added dynamically - Send emails",
                    IsReverted = false
                },
                new PermissionChange
                {
                    Permission = "Calendars.ReadWrite",
                    Timestamp = new DateTime(2026, 2, 5, 9, 28, 42, DateTimeKind.Utc),
                    Reason = "Added dynamically - Read and create calendar events",
                    IsReverted = false
                },
                new PermissionChange
                {
                    Permission = "Team.ReadBasic.All",
                    Timestamp = new DateTime(2026, 2, 5, 9, 30, 15, DateTimeKind.Utc),
                    Reason = "Reverted by user",
                    IsReverted = true
                }
            },
            LastModified = new DateTime(2026, 2, 5, 9, 30, 15, DateTimeKind.Utc)
        };
        _mockPermissionHelper.GetChangeHistory().Returns(history);

        // Act
        var result = await _tools.ViewPermissionHistory();

        // Assert
        Assert.Contains("Permission Change History (3 change(s)):", result);
        Assert.Contains("Last Modified: 2026-02-05 09:30:15 UTC", result);
        Assert.Contains("Active Changes:", result);
        Assert.Contains("Mail.Send", result);
        Assert.Contains("Calendars.ReadWrite", result);
        Assert.Contains("Reverted Changes:", result);
        Assert.Contains("Team.ReadBasic.All", result);
    }

    [Fact]
    public async Task ViewPermissionHistory_WhenExceptionOccurs_ReturnsErrorJson()
    {
        // Arrange
        _mockPermissionHelper.GetChangeHistory().Returns(x => throw new Exception("Test error"));

        // Act
        var result = await _tools.ViewPermissionHistory();

        // Assert
        var json = JsonSerializer.Deserialize<JsonDocument>(result);
        Assert.NotNull(json);
        Assert.True(json.RootElement.TryGetProperty("error", out var errorProp));
        Assert.Equal("Test error", errorProp.GetString());
    }

    [Fact]
    public async Task RevertPermission_WhenSuccessful_ReturnsSuccessMessage()
    {
        // Arrange
        _mockPermissionHelper.RevertPermissionAsync("Mail.Send").Returns(true);

        // Act
        var result = await _tools.RevertPermission("Mail.Send");

        // Assert
        var json = JsonSerializer.Deserialize<JsonDocument>(result);
        Assert.NotNull(json);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean());
        var message = json.RootElement.GetProperty("message").GetString();
        Assert.Contains("Mail.Send", message);
        Assert.Contains("has been reverted", message);
        Assert.Contains("Token cache cleared", message);
    }

    [Fact]
    public async Task RevertPermission_WhenFails_ReturnsFailureMessage()
    {
        // Arrange
        _mockPermissionHelper.RevertPermissionAsync("User.Read").Returns(false);

        // Act
        var result = await _tools.RevertPermission("User.Read");

        // Assert
        var json = JsonSerializer.Deserialize<JsonDocument>(result);
        Assert.NotNull(json);
        Assert.False(json.RootElement.GetProperty("success").GetBoolean());
        var message = json.RootElement.GetProperty("message").GetString();
        Assert.Contains("Failed to revert", message);
        Assert.Contains("User.Read", message);
    }

    [Fact]
    public async Task RevertPermission_WhenExceptionOccurs_ReturnsErrorJson()
    {
        // Arrange
        _mockPermissionHelper.RevertPermissionAsync("Mail.Send")
            .Returns(Task.FromException<bool>(new Exception("Test error")));

        // Act
        var result = await _tools.RevertPermission("Mail.Send");

        // Assert
        var json = JsonSerializer.Deserialize<JsonDocument>(result);
        Assert.NotNull(json);
        Assert.True(json.RootElement.TryGetProperty("error", out var errorProp));
        Assert.Equal("Test error", errorProp.GetString());
    }

    [Fact]
    public async Task RevertAllDynamicPermissions_WhenSuccessful_ReturnsSuccessWithRemainingPermissions()
    {
        // Arrange
        var defaultScopes = new[] { "User.Read", "Mail.Read", "Calendars.Read", "Chat.Read", "Notes.Read" };
        _mockPermissionHelper.RevertAllDynamicPermissionsAsync().Returns(true);
        _mockPermissionHelper.GetDefaultScopes().Returns(defaultScopes);

        // Act
        var result = await _tools.RevertAllDynamicPermissions();

        // Assert
        var json = JsonSerializer.Deserialize<JsonDocument>(result);
        Assert.NotNull(json);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean());
        var message = json.RootElement.GetProperty("message").GetString();
        Assert.Contains("All dynamic permissions have been reverted", message);

        var remainingPermissions = json.RootElement.GetProperty("remainingPermissions");
        var permissionList = new List<string>();
        foreach (var perm in remainingPermissions.EnumerateArray())
        {
            permissionList.Add(perm.GetString()!);
        }
        Assert.Equal(5, permissionList.Count);
        Assert.Contains("User.Read", permissionList);
    }

    [Fact]
    public async Task RevertAllDynamicPermissions_WhenFails_ReturnsFailureMessage()
    {
        // Arrange
        _mockPermissionHelper.RevertAllDynamicPermissionsAsync().Returns(false);

        // Act
        var result = await _tools.RevertAllDynamicPermissions();

        // Assert
        var json = JsonSerializer.Deserialize<JsonDocument>(result);
        Assert.NotNull(json);
        Assert.False(json.RootElement.GetProperty("success").GetBoolean());
        var message = json.RootElement.GetProperty("message").GetString();
        Assert.Contains("Failed to revert dynamic permissions", message);
    }

    [Fact]
    public async Task RevertAllDynamicPermissions_WhenExceptionOccurs_ReturnsErrorJson()
    {
        // Arrange
        _mockPermissionHelper.RevertAllDynamicPermissionsAsync()
            .Returns(Task.FromException<bool>(new Exception("Test error")));

        // Act
        var result = await _tools.RevertAllDynamicPermissions();

        // Assert
        var json = JsonSerializer.Deserialize<JsonDocument>(result);
        Assert.NotNull(json);
        Assert.True(json.RootElement.TryGetProperty("error", out var errorProp));
        Assert.Equal("Test error", errorProp.GetString());
    }

    [Fact]
    public async Task GetDefaultPermissions_ReturnsFormattedDefaultPermissions()
    {
        // Arrange
        var defaultScopes = new[] { "User.Read", "Mail.Read", "Calendars.Read", "Chat.Read", "Notes.Read" };
        _mockPermissionHelper.GetDefaultScopes().Returns(defaultScopes);

        // Act
        var result = await _tools.GetDefaultPermissions();

        // Assert
        Assert.Contains("Default Permissions (always active, cannot be reverted):", result);
        Assert.Contains("User.Read", result);
        Assert.Contains("Mail.Read", result);
        Assert.Contains("Calendars.Read", result);
        Assert.Contains("Chat.Read", result);
        Assert.Contains("Notes.Read", result);
        Assert.Contains("These permissions provide basic read and search capabilities across:", result);
        Assert.Contains("Mail (read and search emails)", result);
        Assert.Contains("Calendar (view events)", result);
        Assert.Contains("Teams (read chats and search messages)", result);
        Assert.Contains("OneNote (read notebooks)", result);
    }

    [Fact]
    public async Task GetDefaultPermissions_WhenExceptionOccurs_ReturnsErrorJson()
    {
        // Arrange
        _mockPermissionHelper.GetDefaultScopes().Returns(x => throw new Exception("Test error"));

        // Act
        var result = await _tools.GetDefaultPermissions();

        // Assert
        var json = JsonSerializer.Deserialize<JsonDocument>(result);
        Assert.NotNull(json);
        Assert.True(json.RootElement.TryGetProperty("error", out var errorProp));
        Assert.Equal("Test error", errorProp.GetString());
    }

    [Fact]
    public async Task Logout_WhenSuccessful_ClearsTokenCacheAndReturnsSuccess()
    {
        // Arrange
        _mockAuthService.ClearTokenCacheAsync().Returns(Task.CompletedTask);

        // Act
        var result = await _tools.Logout();

        // Assert
        await _mockAuthService.Received(1).ClearTokenCacheAsync();
        var json = JsonSerializer.Deserialize<JsonDocument>(result);
        Assert.NotNull(json);
        Assert.True(json.RootElement.GetProperty("success").GetBoolean());
        var message = json.RootElement.GetProperty("message").GetString();
        Assert.Contains("Successfully logged out", message);
        Assert.Contains("cached tokens have been cleared", message);
    }

    [Fact]
    public async Task Logout_WhenExceptionOccurs_ReturnsErrorJson()
    {
        // Arrange
        _mockAuthService.ClearTokenCacheAsync()
            .Returns(Task.FromException(new Exception("Test error")));

        // Act
        var result = await _tools.Logout();

        // Assert
        var json = JsonSerializer.Deserialize<JsonDocument>(result);
        Assert.NotNull(json);
        Assert.True(json.RootElement.TryGetProperty("error", out var errorProp));
        Assert.Equal("Test error", errorProp.GetString());
    }
}
