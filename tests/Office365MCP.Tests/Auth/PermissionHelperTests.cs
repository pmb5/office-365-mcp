using System.Text.Json;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using NSubstitute;
using Office365MCP.Auth;
using Xunit;

namespace Office365MCP.Tests.Auth;

public class PermissionHelperTests : IDisposable
{
    private readonly IGraphAuthService _mockAuthService;
    private readonly IScopeProvider _scopeProvider;
    private readonly ILogger<PermissionHelper> _mockLogger;
    private readonly ILogger<ScopeProvider> _mockScopeLogger;
    private readonly IConfiguration _mockConfiguration;
    private readonly string _testConfigPath;
    private readonly string _testTrackingPath;
    private readonly string _testDirectory;

    public PermissionHelperTests()
    {
        _mockAuthService = Substitute.For<IGraphAuthService>();
        _mockLogger = Substitute.For<ILogger<PermissionHelper>>();
        _mockScopeLogger = Substitute.For<ILogger<ScopeProvider>>();
        _mockConfiguration = Substitute.For<IConfiguration>();

        // Create temporary test directory
        _testDirectory = Path.Combine(Path.GetTempPath(), $"Office365MCP_Tests_{Guid.NewGuid()}");
        Directory.CreateDirectory(_testDirectory);

        _testConfigPath = Path.Combine(_testDirectory, "appsettings.json");
        _testTrackingPath = Path.Combine(_testDirectory, "permission-changes.json");

        // Create a real ScopeProvider configured to use the test config path
        _scopeProvider = new ScopeProvider(_mockScopeLogger);
        var configPathField = typeof(ScopeProvider).GetField("_configPath",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
        configPathField?.SetValue(_scopeProvider, _testConfigPath);
    }

    public void Dispose()
    {
        // Clean up test directory
        if (Directory.Exists(_testDirectory))
        {
            Directory.Delete(_testDirectory, true);
        }
    }

    private PermissionHelper CreatePermissionHelper()
    {
        // Use reflection to create instance and set test paths
        var helper = new PermissionHelper(_mockAuthService, _scopeProvider, _mockLogger);

        var configPathField = typeof(PermissionHelper).GetField("_configPath",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
        configPathField?.SetValue(helper, _testConfigPath);

        var trackingPathField = typeof(PermissionHelper).GetField("_trackingPath",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
        trackingPathField?.SetValue(helper, _testTrackingPath);

        return helper;
    }

    [Fact]
    public void GetDefaultScopes_ReturnsExpectedPermissions()
    {
        // Arrange
        var helper = CreatePermissionHelper();

        // Act
        var scopes = helper.GetDefaultScopes();

        // Assert
        Assert.Equal(5, scopes.Length);
        Assert.Contains("User.Read", scopes);
        Assert.Contains("Mail.Read", scopes);
        Assert.Contains("Calendars.Read", scopes);
        Assert.Contains("Chat.Read", scopes);
        Assert.Contains("Notes.Read", scopes);
    }

    [Fact]
    public void GetCurrentScopes_WhenConfigDoesNotExist_ReturnsDefaults()
    {
        // Arrange
        var helper = CreatePermissionHelper();

        // Act
        var scopes = helper.GetCurrentScopes();

        // Assert
        Assert.Equal(5, scopes.Length);
        Assert.Contains("User.Read", scopes);
    }

    [Fact]
    public void GetCurrentScopes_WhenConfigExists_ReturnsConfiguredScopes()
    {
        // Arrange
        var helper = CreatePermissionHelper();
        var config = new
        {
            Graph = new
            {
                ClientId = "test-client-id",
                TenantId = "organizations",
                Scopes = new[] { "User.Read", "Mail.Read", "Mail.Send", "Calendars.ReadWrite" }
            }
        };
        File.WriteAllText(_testConfigPath, JsonSerializer.Serialize(config, new JsonSerializerOptions { WriteIndented = true }));

        // Act
        var scopes = helper.GetCurrentScopes();

        // Assert
        Assert.Equal(4, scopes.Length);
        Assert.Contains("Mail.Send", scopes);
        Assert.Contains("Calendars.ReadWrite", scopes);
    }

    [Fact]
    public void GetChangeHistory_WhenTrackingFileDoesNotExist_ReturnsEmptyHistory()
    {
        // Arrange
        var helper = CreatePermissionHelper();

        // Act
        var history = helper.GetChangeHistory();

        // Assert
        Assert.NotNull(history);
        Assert.Empty(history.Changes);
    }

    [Fact]
    public void GetChangeHistory_WhenTrackingFileExists_ReturnsHistory()
    {
        // Arrange
        var helper = CreatePermissionHelper();
        var expectedHistory = new PermissionChangeHistory
        {
            Changes = new List<PermissionChange>
            {
                new PermissionChange
                {
                    Permission = "Mail.Send",
                    Timestamp = DateTime.UtcNow,
                    Reason = "Added dynamically - Send emails",
                    AddedBy = "MCP Server",
                    IsReverted = false
                }
            },
            LastModified = DateTime.UtcNow
        };
        File.WriteAllText(_testTrackingPath, JsonSerializer.Serialize(expectedHistory));

        // Act
        var history = helper.GetChangeHistory();

        // Assert
        Assert.NotNull(history);
        Assert.Single(history.Changes);
        Assert.Equal("Mail.Send", history.Changes[0].Permission);
        Assert.Equal("Added dynamically - Send emails", history.Changes[0].Reason);
        Assert.False(history.Changes[0].IsReverted);
    }

    [Fact]
    public async Task RevertPermissionAsync_WhenPermissionIsDefault_ReturnsFalse()
    {
        // Arrange
        var helper = CreatePermissionHelper();
        var config = new
        {
            Graph = new
            {
                ClientId = "test-client-id",
                TenantId = "organizations",
                Scopes = new[] { "User.Read", "Mail.Read", "Calendars.Read" }
            }
        };
        File.WriteAllText(_testConfigPath, JsonSerializer.Serialize(config));

        // Act
        var result = await helper.RevertPermissionAsync("User.Read");

        // Assert
        Assert.False(result);
        await _mockAuthService.DidNotReceive().ClearTokenCacheAsync();
    }

    [Fact]
    public async Task RevertPermissionAsync_WhenPermissionIsDynamic_RemovesAndTracksChange()
    {
        // Arrange
        var helper = CreatePermissionHelper();
        var config = new
        {
            Graph = new
            {
                ClientId = "test-client-id",
                TenantId = "organizations",
                Scopes = new[] { "User.Read", "Mail.Read", "Mail.Send", "Calendars.Read" }
            }
        };
        File.WriteAllText(_testConfigPath, JsonSerializer.Serialize(config));

        // Act
        var result = await helper.RevertPermissionAsync("Mail.Send");

        // Assert
        Assert.True(result);
        await _mockAuthService.Received(1).ClearTokenCacheAsync();

        // Verify permission removed from config
        var updatedConfig = JsonSerializer.Deserialize<JsonDocument>(File.ReadAllText(_testConfigPath));
        var scopes = updatedConfig!.RootElement.GetProperty("Graph").GetProperty("Scopes");
        var scopeList = new List<string>();
        foreach (var scope in scopes.EnumerateArray())
        {
            scopeList.Add(scope.GetString()!);
        }
        Assert.DoesNotContain("Mail.Send", scopeList);

        // Verify tracking file updated
        var history = helper.GetChangeHistory();
        Assert.Single(history.Changes);
        Assert.Equal("Mail.Send", history.Changes[0].Permission);
        Assert.True(history.Changes[0].IsReverted);
        Assert.Contains("Reverted by user", history.Changes[0].Reason);
    }

    [Fact]
    public async Task RevertAllDynamicPermissionsAsync_WhenNoDynamicPermissions_ReturnsTrue()
    {
        // Arrange
        var helper = CreatePermissionHelper();
        var config = new
        {
            Graph = new
            {
                ClientId = "test-client-id",
                TenantId = "organizations",
                Scopes = new[] { "User.Read", "Mail.Read", "Calendars.Read", "Chat.Read", "Notes.Read" }
            }
        };
        File.WriteAllText(_testConfigPath, JsonSerializer.Serialize(config));

        // Act
        var result = await helper.RevertAllDynamicPermissionsAsync();

        // Assert
        Assert.True(result);
        await _mockAuthService.DidNotReceive().ClearTokenCacheAsync();
    }

    [Fact]
    public async Task RevertAllDynamicPermissionsAsync_WhenDynamicPermissionsExist_RemovesAll()
    {
        // Arrange
        var helper = CreatePermissionHelper();
        var config = new
        {
            Graph = new
            {
                ClientId = "test-client-id",
                TenantId = "organizations",
                Scopes = new[] { "User.Read", "Mail.Read", "Calendars.Read", "Chat.Read", "Notes.Read", "Mail.Send", "Calendars.ReadWrite" }
            }
        };
        File.WriteAllText(_testConfigPath, JsonSerializer.Serialize(config));

        // Act
        var result = await helper.RevertAllDynamicPermissionsAsync();

        // Assert
        Assert.True(result);
        await _mockAuthService.Received(1).ClearTokenCacheAsync();

        // Verify only default permissions remain
        var currentScopes = helper.GetCurrentScopes();
        Assert.Equal(5, currentScopes.Length);
        Assert.DoesNotContain("Mail.Send", currentScopes);
        Assert.DoesNotContain("Calendars.ReadWrite", currentScopes);
        Assert.Contains("Calendars.Read", currentScopes);

        // Verify tracking
        var history = helper.GetChangeHistory();
        Assert.Equal(2, history.Changes.Count);
        Assert.All(history.Changes, change => Assert.True(change.IsReverted));
    }

    [Fact]
    public async Task HandlePermissionErrorAsync_WhenNoPermissionsDetected_ReturnsFalse()
    {
        // Arrange
        var helper = CreatePermissionHelper();
        var errorResponse = "{\"error\": {\"message\": \"Some generic error\"}}";

        // Act
        var result = await helper.HandlePermissionErrorAsync(errorResponse);

        // Assert
        Assert.False(result);
    }

    [Fact]
    public void GetCurrentScopes_WhenConfigHasNoScopes_ReturnsDefaults()
    {
        // Arrange
        var helper = CreatePermissionHelper();
        var config = new
        {
            Graph = new
            {
                ClientId = "test-client-id",
                TenantId = "organizations"
            }
        };
        File.WriteAllText(_testConfigPath, JsonSerializer.Serialize(config));

        // Act
        var scopes = helper.GetCurrentScopes();

        // Assert
        Assert.Equal(5, scopes.Length);
        Assert.Contains("User.Read", scopes);
    }

    [Fact]
    public async Task RevertPermissionAsync_WhenConfigDoesNotExist_ReturnsTrue()
    {
        // Arrange
        var helper = CreatePermissionHelper();

        // Act
        var result = await helper.RevertPermissionAsync("Mail.Send");

        // Assert - Returns true because it's not a default permission and no error occurred
        // The permission removal is a no-op when config doesn't exist, but it still tracks the change
        Assert.True(result);
        await _mockAuthService.Received(1).ClearTokenCacheAsync();
    }

    [Fact]
    public void GetChangeHistory_WhenTrackingFileIsCorrupt_ReturnsEmptyHistory()
    {
        // Arrange
        var helper = CreatePermissionHelper();
        File.WriteAllText(_testTrackingPath, "{ invalid json }");

        // Act
        var history = helper.GetChangeHistory();

        // Assert
        Assert.NotNull(history);
        Assert.Empty(history.Changes);
    }
}
