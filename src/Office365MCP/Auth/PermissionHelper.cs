using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Extensions.Logging;

namespace Office365MCP.Auth;

public interface IPermissionHelper
{
    Task<bool> HandlePermissionErrorAsync(string errorResponse);
    string[] GetDefaultScopes();
    string[] GetCurrentScopes();
    PermissionChangeHistory GetChangeHistory();
    Task<bool> RevertPermissionAsync(string permission);
    Task<bool> RevertAllDynamicPermissionsAsync();
}

public class PermissionHelper : IPermissionHelper
{
    private readonly IGraphAuthService _authService;
    private readonly IScopeProvider _scopeProvider;
    private readonly ILogger<PermissionHelper> _logger;
    private readonly string _configPath;
    private readonly string _trackingPath;

    // Default permissions granted on first startup (basic read access across all services)
    private static readonly string[] DefaultPermissions = new[]
    {
        "User.Read",           // Required for Graph API access
        "Mail.Read",           // Read emails and search
        "Calendars.Read",      // View calendar events
        "Chat.Read",           // Read Teams chats and search messages
        "Notes.Read"           // Read OneNote notebooks
    };

    // Known Graph API permissions we support
    private static readonly HashSet<string> SupportedPermissions = new()
    {
        "User.Read",
        "Mail.Read",
        "Mail.Send",
        "Mail.ReadBasic",
        "Mail.ReadWrite",
        "Calendars.Read",
        "Calendars.ReadWrite",
        "Team.ReadBasic.All",
        "Channel.ReadBasic.All",
        "ChannelMessage.Read.All",
        "ChannelMessage.Send",
        "Chat.Read",
        "Chat.ReadWrite",
        "Notes.Read",
        "Notes.ReadWrite",
        "Notes.Create"
    };

    // Permission descriptions
    private static readonly Dictionary<string, string> PermissionDescriptions = new()
    {
        ["User.Read"] = "Read user profile",
        ["Mail.Read"] = "Read email messages",
        ["Mail.Send"] = "Send emails",
        ["Mail.ReadBasic"] = "Read basic email properties",
        ["Mail.ReadWrite"] = "Read and write email",
        ["Calendars.Read"] = "Read calendar events",
        ["Calendars.ReadWrite"] = "Read and create calendar events",
        ["Team.ReadBasic.All"] = "Read basic team information",
        ["Channel.ReadBasic.All"] = "Read basic channel information",
        ["ChannelMessage.Read.All"] = "Read channel messages",
        ["ChannelMessage.Send"] = "Send channel messages",
        ["Chat.Read"] = "Read chat messages",
        ["Chat.ReadWrite"] = "Read and send chat messages",
        ["Notes.Read"] = "Read OneNote notebooks",
        ["Notes.ReadWrite"] = "Read and write OneNote",
        ["Notes.Create"] = "Create OneNote pages"
    };

    public PermissionHelper(IGraphAuthService authService, IScopeProvider scopeProvider, ILogger<PermissionHelper> logger)
    {
        _authService = authService;
        _scopeProvider = scopeProvider;
        _logger = logger;

        var basePath = AppContext.BaseDirectory;
        _configPath = Path.Combine(basePath, "appsettings.json");

        // Store tracking file in the same location as token cache
        var trackingFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Office365MCP");
        Directory.CreateDirectory(trackingFolder);
        _trackingPath = Path.Combine(trackingFolder, "permission-changes.json");
    }

    public string[] GetDefaultScopes()
    {
        return _scopeProvider.GetDefaultScopes();
    }

    public string[] GetCurrentScopes()
    {
        return _scopeProvider.GetCurrentScopes();
    }

    public PermissionChangeHistory GetChangeHistory()
    {
        try
        {
            if (!File.Exists(_trackingPath))
            {
                return new PermissionChangeHistory();
            }

            var json = File.ReadAllText(_trackingPath);
            var history = JsonSerializer.Deserialize<PermissionChangeHistory>(json);
            return history ?? new PermissionChangeHistory();
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to read permission change history");
            return new PermissionChangeHistory();
        }
    }

    public async Task<bool> RevertPermissionAsync(string permission)
    {
        try
        {
            // Check if permission is in defaults (can't revert defaults)
            if (DefaultPermissions.Contains(permission))
            {
                _logger.LogWarning("Cannot revert default permission: {Permission}", permission);
                await Console.Error.WriteLineAsync($"  ✗ Cannot revert '{permission}' - it's a default permission");
                return false;
            }

            // Remove from config
            await RemovePermissionFromConfigAsync(permission);

            // Mark as reverted in tracking
            await TrackPermissionChange(permission, "Reverted by user", isRevert: true);

            // Clear token cache
            await _authService.ClearTokenCacheAsync();

            _logger.LogInformation("Reverted permission {Permission}", permission);
            await Console.Error.WriteLineAsync($"  ✓ Permission '{permission}' reverted");
            await Console.Error.WriteLineAsync("  ✓ Token cache cleared");
            await Console.Error.WriteLineAsync("  → Please restart any operations. New tokens will be acquired without this permission.");

            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to revert permission {Permission}", permission);
            await Console.Error.WriteLineAsync($"  ✗ Error reverting permission: {ex.Message}");
            return false;
        }
    }

    public async Task<bool> RevertAllDynamicPermissionsAsync()
    {
        try
        {
            var currentScopes = GetCurrentScopes();
            var dynamicPermissions = currentScopes.Except(DefaultPermissions).ToList();

            if (dynamicPermissions.Count == 0)
            {
                await Console.Error.WriteLineAsync("  ℹ No dynamic permissions to revert. Only default permissions are active.");
                return true;
            }

            await Console.Error.WriteLineAsync($"  Found {dynamicPermissions.Count} dynamic permission(s) to revert:");
            foreach (var perm in dynamicPermissions)
            {
                await Console.Error.WriteLineAsync($"    - {perm}");
            }

            // Remove all dynamic permissions from config
            foreach (var permission in dynamicPermissions)
            {
                await RemovePermissionFromConfigAsync(permission);
                await TrackPermissionChange(permission, "Reverted with all dynamic permissions", isRevert: true);
            }

            // Clear token cache
            await _authService.ClearTokenCacheAsync();

            _logger.LogInformation("Reverted {Count} dynamic permissions", dynamicPermissions.Count);
            await Console.Error.WriteLineAsync("");
            await Console.Error.WriteLineAsync($"  ✓ Reverted {dynamicPermissions.Count} permission(s)");
            await Console.Error.WriteLineAsync("  ✓ Token cache cleared");
            await Console.Error.WriteLineAsync("  → Server reset to default permissions only");

            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to revert dynamic permissions");
            await Console.Error.WriteLineAsync($"  ✗ Error: {ex.Message}");
            return false;
        }
    }

    private async Task TrackPermissionChange(string permission, string reason, bool isRevert = false)
    {
        try
        {
            var history = GetChangeHistory();

            var change = new PermissionChange
            {
                Permission = permission,
                Timestamp = DateTime.UtcNow,
                Reason = reason,
                AddedBy = "MCP Server",
                IsReverted = isRevert
            };

            history.Changes.Add(change);
            history.LastModified = DateTime.UtcNow;

            var options = new JsonSerializerOptions { WriteIndented = true };
            var json = JsonSerializer.Serialize(history, options);
            await File.WriteAllTextAsync(_trackingPath, json);

            _logger.LogInformation("Tracked permission change: {Permission} - {Reason}", permission, reason);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to track permission change for {Permission}", permission);
        }
    }

    private async Task RemovePermissionFromConfigAsync(string permission)
    {
        if (!File.Exists(_configPath))
        {
            return;
        }

        var json = await File.ReadAllTextAsync(_configPath);
        using var doc = JsonDocument.Parse(json);

        var config = new Dictionary<string, object>();

        foreach (var property in doc.RootElement.EnumerateObject())
        {
            if (property.Name == "Graph")
            {
                var graphConfig = new Dictionary<string, object>();

                foreach (var graphProp in property.Value.EnumerateObject())
                {
                    if (graphProp.Name == "Scopes")
                    {
                        var scopes = new List<string>();
                        foreach (var scope in graphProp.Value.EnumerateArray())
                        {
                            var scopeValue = scope.GetString();
                            if (!string.IsNullOrEmpty(scopeValue) &&
                                !scopeValue.Equals(permission, StringComparison.OrdinalIgnoreCase))
                            {
                                scopes.Add(scopeValue);
                            }
                        }

                        if (scopes.Count > 0)
                        {
                            graphConfig["Scopes"] = scopes;
                        }
                    }
                    else
                    {
                        var value = graphProp.Value.ValueKind == JsonValueKind.String
                            ? graphProp.Value.GetString()!
                            : graphProp.Value.ToString();
                        graphConfig[graphProp.Name] = value;
                    }
                }

                config["Graph"] = graphConfig;
            }
            else
            {
                var value = property.Value.ValueKind == JsonValueKind.String
                    ? property.Value.GetString()!
                    : property.Value.ToString();
                config[property.Name] = value;
            }
        }

        var options = new JsonSerializerOptions { WriteIndented = true };
        var updatedJson = JsonSerializer.Serialize(config, options);
        await File.WriteAllTextAsync(_configPath, updatedJson);

        _logger.LogInformation("Removed permission {Permission} from config", permission);
    }

    public async Task<bool> HandlePermissionErrorAsync(string errorResponse)
    {
        try
        {
            // Try to parse the error response to extract the missing permission
            var missingPermissions = ExtractMissingPermissions(errorResponse);

            if (missingPermissions.Count == 0)
            {
                _logger.LogWarning("Could not determine missing permission from error response");
                return false;
            }

            // Prompt for each missing permission
            foreach (var permission in missingPermissions)
            {
                if (!SupportedPermissions.Contains(permission))
                {
                    _logger.LogWarning("Permission {Permission} is not in our supported list", permission);
                    continue;
                }

                if (await PromptAndAddPermissionAsync(permission))
                {
                    return true; // Permission added successfully
                }
            }

            return false;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error handling permission error");
            return false;
        }
    }

    private static List<string> ExtractMissingPermissions(string errorResponse)
    {
        var permissions = new List<string>();

        try
        {
            // Parse JSON error response
            var jsonDoc = JsonDocument.Parse(errorResponse);
            var root = jsonDoc.RootElement;

            // Check for error message
            if (root.TryGetProperty("error", out var error))
            {
                if (error.TryGetProperty("message", out var message))
                {
                    var msg = message.GetString() ?? "";

                    // Look for permission names in the error message
                    // Common patterns: "Requires permission: Mail.Read" or "scope 'Mail.Read'"
                    foreach (var permission in SupportedPermissions)
                    {
                        if (msg.Contains(permission, StringComparison.OrdinalIgnoreCase))
                        {
                            permissions.Add(permission);
                        }
                    }
                }
            }
        }
        catch
        {
            // If JSON parsing fails, try regex patterns
            foreach (var permission in SupportedPermissions)
            {
                if (errorResponse.Contains(permission, StringComparison.OrdinalIgnoreCase))
                {
                    permissions.Add(permission);
                }
            }
        }

        return permissions.Distinct().ToList();
    }

    private async Task<bool> PromptAndAddPermissionAsync(string permission)
    {
        var description = PermissionDescriptions.TryGetValue(permission, out var desc)
            ? desc
            : "Access Microsoft 365 resources";

        // Display permission request to user via stderr
        await Console.Error.WriteLineAsync("");
        await Console.Error.WriteLineAsync("╔═══════════════════════════════════════════════════════════════╗");
        await Console.Error.WriteLineAsync("║                  PERMISSION REQUIRED                          ║");
        await Console.Error.WriteLineAsync("╚═══════════════════════════════════════════════════════════════╝");
        await Console.Error.WriteLineAsync("");
        await Console.Error.WriteLineAsync($"  The operation you requested requires the following permission:");
        await Console.Error.WriteLineAsync("");
        await Console.Error.WriteLineAsync($"    Permission: {permission}");
        await Console.Error.WriteLineAsync($"    Purpose:    {description}");
        await Console.Error.WriteLineAsync("");
        await Console.Error.WriteLineAsync("  Would you like to add this permission to your configuration?");
        await Console.Error.WriteLineAsync("  (A browser window will open to grant consent)");
        await Console.Error.WriteLineAsync("");
        await Console.Error.WriteLineAsync("  Type 'yes' to add permission, or 'no' to cancel:");
        Console.Error.Write("  > ");

        var response = Console.ReadLine()?.Trim().ToLowerInvariant();

        if (response != "yes" && response != "y")
        {
            await Console.Error.WriteLineAsync("");
            await Console.Error.WriteLineAsync("  ✗ Permission not added. Operation cancelled.");
            await Console.Error.WriteLineAsync("");
            return false;
        }

        try
        {
            // Add permission to config
            await AddPermissionToConfigAsync(permission);

            // Track this permission change
            await TrackPermissionChange(permission, $"Added dynamically - {description}");

            // Clear token cache to force re-authentication
            await _authService.ClearTokenCacheAsync();

            await Console.Error.WriteLineAsync("");
            await Console.Error.WriteLineAsync("  ✓ Permission added to configuration");
            await Console.Error.WriteLineAsync("  ✓ Change tracked for future reference");
            await Console.Error.WriteLineAsync("  ✓ Token cache cleared");
            await Console.Error.WriteLineAsync("  → Please retry your operation. A browser will open for consent.");
            await Console.Error.WriteLineAsync("");

            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to add permission {Permission}", permission);
            await Console.Error.WriteLineAsync("");
            await Console.Error.WriteLineAsync($"  ✗ Error adding permission: {ex.Message}");
            await Console.Error.WriteLineAsync("");
            return false;
        }
    }

    private async Task AddPermissionToConfigAsync(string permission)
    {
        if (!File.Exists(_configPath))
        {
            throw new FileNotFoundException($"Configuration file not found: {_configPath}");
        }

        // Read current config
        var json = await File.ReadAllTextAsync(_configPath);
        using var doc = JsonDocument.Parse(json);

        // Build new config with added scope
        var config = new Dictionary<string, object>();

        foreach (var property in doc.RootElement.EnumerateObject())
        {
            if (property.Name == "Graph")
            {
                var graphConfig = new Dictionary<string, object>();

                foreach (var graphProp in property.Value.EnumerateObject())
                {
                    if (graphProp.Name == "Scopes")
                    {
                        var scopes = new List<string>();
                        foreach (var scope in graphProp.Value.EnumerateArray())
                        {
                            var scopeValue = scope.GetString();
                            if (!string.IsNullOrEmpty(scopeValue))
                            {
                                scopes.Add(scopeValue);
                            }
                        }

                        // Add new permission if not already present
                        if (!scopes.Contains(permission, StringComparer.OrdinalIgnoreCase))
                        {
                            scopes.Add(permission);
                            _logger.LogInformation("Added permission {Permission} to configuration", permission);
                        }

                        graphConfig["Scopes"] = scopes;
                    }
                    else
                    {
                        var value = graphProp.Value.ValueKind == JsonValueKind.String
                            ? graphProp.Value.GetString()!
                            : graphProp.Value.ToString();
                        graphConfig[graphProp.Name] = value;
                    }
                }

                config["Graph"] = graphConfig;
            }
            else
            {
                // Copy other top-level properties as-is
                var value = property.Value.ValueKind == JsonValueKind.String
                    ? property.Value.GetString()!
                    : property.Value.ToString();
                config[property.Name] = value;
            }
        }

        // Write back to file
        var options = new JsonSerializerOptions { WriteIndented = true };
        var updatedJson = JsonSerializer.Serialize(config, options);
        await File.WriteAllTextAsync(_configPath, updatedJson);
    }
}
