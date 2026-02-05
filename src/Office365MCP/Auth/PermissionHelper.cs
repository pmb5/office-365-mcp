using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Extensions.Logging;

namespace Office365MCP.Auth;

public interface IPermissionHelper
{
    Task<bool> HandlePermissionErrorAsync(string errorResponse);
    string[] GetDefaultScopes();
    string[] GetCurrentScopes();
}

public class PermissionHelper : IPermissionHelper
{
    private readonly IGraphAuthService _authService;
    private readonly ILogger<PermissionHelper> _logger;
    private readonly string _configPath;

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

    public PermissionHelper(IGraphAuthService authService, ILogger<PermissionHelper> logger)
    {
        _authService = authService;
        _logger = logger;

        var basePath = AppContext.BaseDirectory;
        _configPath = Path.Combine(basePath, "appsettings.json");
    }

    public string[] GetDefaultScopes()
    {
        return DefaultPermissions;
    }

    public string[] GetCurrentScopes()
    {
        try
        {
            if (!File.Exists(_configPath))
            {
                return DefaultPermissions;
            }

            var json = File.ReadAllText(_configPath);
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            if (root.TryGetProperty("Graph", out var graphElement) &&
                graphElement.TryGetProperty("Scopes", out var scopesElement))
            {
                var scopes = new List<string>();
                foreach (var scope in scopesElement.EnumerateArray())
                {
                    var scopeValue = scope.GetString();
                    if (!string.IsNullOrEmpty(scopeValue))
                    {
                        scopes.Add(scopeValue);
                    }
                }

                return scopes.Count > 0 ? scopes.ToArray() : DefaultPermissions;
            }

            return DefaultPermissions;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to read scopes from config, using defaults");
            return DefaultPermissions;
        }
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

            // Clear token cache to force re-authentication
            await _authService.ClearTokenCacheAsync();

            await Console.Error.WriteLineAsync("");
            await Console.Error.WriteLineAsync("  ✓ Permission added to configuration");
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
