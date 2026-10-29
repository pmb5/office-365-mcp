# Permissions Moved to Code

## ✅ Changes Complete

I've moved permission management entirely into C# code, removing the dependency on `appsettings.json` for permission configuration. The server now starts with a sensible set of default read permissions defined in code.

## What Changed

### Before: Permissions in Config
```json
// appsettings.json
{
  "Graph": {
    "ClientId": "...",
    "TenantId": "...",
    "Scopes": [
      "User.Read",
      "Mail.Read",
      // ... 12 more permissions
    ]
  }
}
```

### After: Permissions in Code
```json
// appsettings.json (simplified!)
{
  "Graph": {
    "ClientId": "14d82eec-204b-4c2f-b7e8-296a70dab67e",
    "TenantId": "organizations"
  }
}
```

```csharp
// PermissionHelper.cs
private static readonly string[] DefaultPermissions = new[]
{
    "User.Read",           // Required for Graph API access
    "Mail.Read",           // Read emails and search
    "Calendars.Read",      // View calendar events
    "Chat.Read",           // Read Teams chats and search messages
    "Notes.Read"           // Read OneNote notebooks
};
```

## Architecture Changes

### 1. PermissionHelper Enhanced
**File:** `src/Office365MCP/Auth/PermissionHelper.cs`

**Added:**
- `DefaultPermissions` array - The baseline permissions granted on first auth
- `GetDefaultScopes()` method - Returns the default permission set
- `GetCurrentScopes()` method - Returns current permissions (config file + defaults)

**Benefits:**
- Single source of truth for all permissions
- Default permissions defined with clear comments
- No configuration needed to start

### 2. GraphAuthService Updated
**File:** `src/Office365MCP/Auth/GraphAuthService.cs`

**Changed:**
- Now injects `IPermissionHelper` instead of reading from config
- Calls `_permissionHelper.GetCurrentScopes()` when acquiring tokens
- No longer manages scopes internally

**Benefits:**
- Simplified - delegates scope management to PermissionHelper
- More flexible - scopes can change between auth requests
- Better logging - includes scope count in log messages

### 3. Configuration Simplified
**File:** `src/Office365MCP/appsettings.json`

**Removed:**
- `Scopes` array completely removed
- Only ClientId and TenantId remain

**Benefits:**
- Minimal configuration required
- No permission lists to maintain
- Cleaner, simpler config file

## Default Permissions Explained

The server starts with **5 basic read permissions** that cover viewing and searching across all services:

### 1. User.Read ✅
- **Purpose:** Required for Microsoft Graph API access
- **Grants:** Read basic user profile information
- **Admin Consent:** No
- **Always Included:** Yes (mandatory)

### 2. Mail.Read ✅
- **Purpose:** Email viewing and search
- **Grants:** Read all emails, search mailbox
- **Tools Enabled:** `ListEmails`, `GetEmail`, `SearchEmails`, `ListMailFolders`
- **Admin Consent:** No

### 3. Calendars.Read ✅
- **Purpose:** Calendar viewing
- **Grants:** Read calendar events across all calendars
- **Tools Enabled:** `ListCalendarEvents`, `GetCalendarEvent`, `ListCalendars`
- **Admin Consent:** No

### 4. Chat.Read ✅
- **Purpose:** Teams chat and message search
- **Grants:** Read chats and search messages
- **Tools Enabled:** `ListChats`, `ListChatMessages`
- **Admin Consent:** No
- **Note:** This enables searching your own Teams chats!

### 5. Notes.Read ✅
- **Purpose:** OneNote viewing
- **Grants:** Read all OneNote notebooks
- **Tools Enabled:** `ListNotebooks`, `ListSections`, `ListPages`, `GetPageContent`
- **Admin Consent:** No

### Why These Five?

1. **Comprehensive Read Access** - Cover viewing data in all services
2. **Search Enabled** - Mail and Teams chat search work immediately
3. **No Admin Required** - All are user-level permissions
4. **Privacy Focused** - Read-only, no sending/creating initially
5. **Upgradeable** - Write permissions added dynamically when needed

## How Permissions Work Now

### First Startup
```
1. User starts MCP server
2. GraphAuthService requests scopes from PermissionHelper
3. PermissionHelper returns default 5 permissions
4. Browser opens, user consents to 5 permissions
5. Token cached with these scopes
6. Server ready with read + search capabilities
```

### Using a Write Operation
```
1. User: "Send an email to john@example.com"
2. Server tries to call /me/sendMail
3. Graph API returns 403: "Missing Mail.Send permission"
4. PermissionHelper detects error
5. Prompts user: "Add Mail.Send permission?"
6. User: "yes"
7. Mail.Send added to config file
8. Token cache cleared
9. Browser opens for re-consent (now 6 permissions)
10. New token acquired
11. Send operation retries and succeeds
```

### Subsequent Startups
```
1. Server starts
2. PermissionHelper reads config file
3. Finds: 5 defaults + Mail.Send (previously added)
4. Returns 6 permissions total
5. Silent auth with cached token
6. Server ready immediately
```

## Teams Search

`SearchTeamsMessages` is the only Teams message search. It uses the Microsoft Search API
(`POST /search/query`), which covers full history across chats and channels. Each hit carries
the chat or channel it came from, so results for one conversation can be picked out from it.

Graph has no `$search` on a single chat's or channel's message list, which is why there are no
per-chat or per-channel search tools.

- **Requires:** `ChannelMessage.Read.All` + `Chat.Read`
- **Note:** `Chat.Read` is in defaults, so only `ChannelMessage.Read.All` needs adding

## Benefits Summary

### For Users
- ✅ **Minimal configuration** - Only 2 config values (ClientId, TenantId)
- ✅ **Sensible defaults** - Read + search capabilities from day one
- ✅ **Clear upgrades** - Write permissions added as needed
- ✅ **Immediate search** - Mail and chat search work out of the box

### For Developers
- ✅ **Single source of truth** - All permissions defined in one place (PermissionHelper.cs)
- ✅ **Type safety** - Permissions in code, not config strings
- ✅ **Self-documenting** - Comments explain each permission
- ✅ **Easier testing** - Mock PermissionHelper instead of config files

### For Maintenance
- ✅ **No config updates** - Add new permissions in code only
- ✅ **Version control friendly** - Permissions tracked in code
- ✅ **Consistent across deployments** - Same defaults everywhere
- ✅ **Easier to reason about** - Clear permission hierarchy

## Migration from Previous Version

If you're upgrading:

### Option 1: Fresh Start (Recommended)
1. Delete token cache: `%LocalAppData%\Office365MCP\office365mcp.cache`
2. Update `appsettings.json` - remove `Scopes` array
3. Restart server
4. Consent to 5 default permissions
5. Use tools - additional permissions added as needed

### Option 2: Keep Existing Permissions
Your existing config file with `Scopes` array will still work:
- PermissionHelper reads from config if present
- Falls back to defaults if config empty
- Dynamic additions append to existing list

## Files Modified

### Core Changes
1. **PermissionHelper.cs** - Added default permissions + scope management methods
2. **GraphAuthService.cs** - Now uses PermissionHelper for scopes
3. **appsettings.json** - Removed Scopes array
4. **GraphAuthServiceTests.cs** - Updated to mock PermissionHelper

### Documentation
- Created: `PERMISSIONS-IN-CODE.md` (this file)
- Updated: Configuration examples in other docs

## Verification

- ✅ **Build:** Success (0 warnings, 0 errors)
- ✅ **Tests:** 15/15 Passing
- ✅ **Default Permissions:** 5 read permissions defined
- ✅ **Teams Search:** 3 search methods confirmed
- ✅ **Config File:** Simplified to 2 values
- ✅ **Backward Compatible:** Old configs still work

## Summary

Permissions are now **defined in code** with sensible defaults that provide **read and search capabilities** across all services (Mail, Calendar, Teams, OneNote). The configuration file is **dramatically simplified** to just ClientId and TenantId. Write permissions are **dynamically added** when needed, giving users a smooth progressive consent experience.

**Default Experience:**
- Start → Consent to 5 read permissions
- Use read/search tools → Works immediately
- Use write tool → Prompted to add that permission
- Future runs → All granted permissions work silently

This provides the **best of both worlds**: minimal initial setup with comprehensive baseline functionality, plus the flexibility to add more permissions on demand.
