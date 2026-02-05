using System.ComponentModel;
using System.Text;
using System.Text.Json;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Server;
using Office365MCP.Auth;

namespace Office365MCP.Tools;

[McpServerToolType]
public class PermissionManagementTools
{
    private readonly IPermissionHelper _permissionHelper;
    private readonly ILogger<PermissionManagementTools> _logger;

    public PermissionManagementTools(IPermissionHelper permissionHelper, ILogger<PermissionManagementTools> logger)
    {
        _permissionHelper = permissionHelper;
        _logger = logger;
    }

    [McpServerTool, Description("List all currently active permissions (default + dynamically added)")]
    public async Task<string> ListCurrentPermissions()
    {
        try
        {
            var defaultScopes = _permissionHelper.GetDefaultScopes();
            var currentScopes = _permissionHelper.GetCurrentScopes();
            var dynamicScopes = currentScopes.Except(defaultScopes).ToList();

            var result = new StringBuilder();
            result.AppendLine("Current Permissions:");
            result.AppendLine();

            result.AppendLine("Default Permissions (cannot be reverted):");
            foreach (var scope in defaultScopes)
            {
                result.AppendLine($"  ✓ {scope}");
            }

            if (dynamicScopes.Count > 0)
            {
                result.AppendLine();
                result.AppendLine("Dynamically Added Permissions (can be reverted):");
                foreach (var scope in dynamicScopes)
                {
                    result.AppendLine($"  + {scope}");
                }
            }
            else
            {
                result.AppendLine();
                result.AppendLine("No dynamic permissions added yet.");
            }

            result.AppendLine();
            result.AppendLine($"Total: {currentScopes.Length} permission(s) active");

            return result.ToString();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to list current permissions");
            return JsonSerializer.Serialize(new { error = ex.Message });
        }
    }

    [McpServerTool, Description("View the history of all permission changes made by the MCP server")]
    public async Task<string> ViewPermissionHistory()
    {
        try
        {
            var history = _permissionHelper.GetChangeHistory();

            if (history.Changes.Count == 0)
            {
                return "No permission changes recorded yet.";
            }

            var result = new StringBuilder();
            result.AppendLine($"Permission Change History ({history.Changes.Count} change(s)):");
            result.AppendLine($"Last Modified: {history.LastModified:yyyy-MM-dd HH:mm:ss} UTC");
            result.AppendLine();

            var activeChanges = history.Changes.Where(c => !c.IsReverted).ToList();
            var revertedChanges = history.Changes.Where(c => c.IsReverted).ToList();

            if (activeChanges.Count > 0)
            {
                result.AppendLine("Active Changes:");
                foreach (var change in activeChanges.OrderByDescending(c => c.Timestamp))
                {
                    result.AppendLine($"  [{change.Timestamp:yyyy-MM-dd HH:mm:ss}] + {change.Permission}");
                    result.AppendLine($"    Reason: {change.Reason}");
                    result.AppendLine();
                }
            }

            if (revertedChanges.Count > 0)
            {
                result.AppendLine("Reverted Changes:");
                foreach (var change in revertedChanges.OrderByDescending(c => c.Timestamp))
                {
                    result.AppendLine($"  [{change.Timestamp:yyyy-MM-dd HH:mm:ss}] - {change.Permission}");
                    result.AppendLine($"    Reason: {change.Reason}");
                    result.AppendLine();
                }
            }

            return result.ToString();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to view permission history");
            return JsonSerializer.Serialize(new { error = ex.Message });
        }
    }

    [McpServerTool, Description("Revert a specific dynamically added permission")]
    public async Task<string> RevertPermission(
        [Description("The permission to revert (e.g., 'Mail.Send')")] string permission)
    {
        try
        {
            var success = await _permissionHelper.RevertPermissionAsync(permission);

            if (success)
            {
                return JsonSerializer.Serialize(new
                {
                    success = true,
                    message = $"Permission '{permission}' has been reverted. Token cache cleared. The permission will no longer be included in future authentication requests."
                });
            }

            return JsonSerializer.Serialize(new
            {
                success = false,
                message = $"Failed to revert permission '{permission}'. It may be a default permission or not currently active."
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to revert permission {Permission}", permission);
            return JsonSerializer.Serialize(new { error = ex.Message });
        }
    }

    [McpServerTool, Description("Revert all dynamically added permissions, keeping only the default permissions")]
    public async Task<string> RevertAllDynamicPermissions()
    {
        try
        {
            var success = await _permissionHelper.RevertAllDynamicPermissionsAsync();

            if (success)
            {
                var defaultScopes = _permissionHelper.GetDefaultScopes();
                return JsonSerializer.Serialize(new
                {
                    success = true,
                    message = "All dynamic permissions have been reverted. Server reset to default permissions only.",
                    remainingPermissions = defaultScopes
                });
            }

            return JsonSerializer.Serialize(new
            {
                success = false,
                message = "Failed to revert dynamic permissions."
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to revert all dynamic permissions");
            return JsonSerializer.Serialize(new { error = ex.Message });
        }
    }

    [McpServerTool, Description("Get the list of default permissions that are always granted")]
    public async Task<string> GetDefaultPermissions()
    {
        try
        {
            var defaultScopes = _permissionHelper.GetDefaultScopes();

            var result = new StringBuilder();
            result.AppendLine("Default Permissions (always active, cannot be reverted):");
            result.AppendLine();

            foreach (var scope in defaultScopes)
            {
                result.AppendLine($"  ✓ {scope}");
            }

            result.AppendLine();
            result.AppendLine("These permissions provide basic read and search capabilities across:");
            result.AppendLine("  - Mail (read and search emails)");
            result.AppendLine("  - Calendar (view events)");
            result.AppendLine("  - Teams (read chats and search messages)");
            result.AppendLine("  - OneNote (read notebooks)");

            return result.ToString();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to get default permissions");
            return JsonSerializer.Serialize(new { error = ex.Message });
        }
    }
}
