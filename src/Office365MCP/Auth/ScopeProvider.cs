using System.Text.Json;
using Microsoft.Extensions.Logging;

namespace Office365MCP.Auth;

public class ScopeProvider : IScopeProvider
{
    private readonly ILogger<ScopeProvider> _logger;
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

    public ScopeProvider(ILogger<ScopeProvider> logger)
    {
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
}
