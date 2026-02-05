# Configuration Simplification

## ✅ Changes Complete

I've simplified the configuration by removing the static permission list from `appsettings.json`. The MCP server now starts with minimal configuration and adds permissions dynamically as needed.

## What Changed

### Before
```json
{
  "Graph": {
    "ClientId": "14d82eec-204b-4c2f-b7e8-296a70dab67e",
    "TenantId": "organizations",
    "Scopes": [
      "User.Read",
      "Mail.Read",
      "Mail.Send",
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
    ]
  }
}
```

### After
```json
{
  "Graph": {
    "ClientId": "14d82eec-204b-4c2f-b7e8-296a70dab67e",
    "TenantId": "organizations",
    "Scopes": [
      "User.Read"
    ]
  }
}
```

## Key Improvements

### 1. Minimal Initial Configuration
- Only `User.Read` is configured by default
- No need to pre-configure all possible permissions
- Cleaner, simpler configuration file

### 2. Smart Defaults
- `GraphAuthService` automatically ensures `User.Read` is always included
- If `Scopes` array is empty or missing, defaults to `["User.Read"]`
- Prevents configuration errors

### 3. Dynamic Permission Management
- Permissions are defined in C# code (`PermissionHelper.cs`)
- Single source of truth for supported permissions
- No duplication between code and config

### 4. Just-In-Time Permission Granting
- Start with minimal permissions
- Add permissions only when actually needed
- Follows principle of least privilege

## How It Works Now

### Startup
```
1. Server starts with User.Read permission
2. User makes first request (e.g., "List my emails")
3. Server detects missing Mail.Read permission
4. User is prompted to add Mail.Read
5. User consents, permission added to config
6. Future requests work without prompting
```

### Permission Storage
```
Supported Permissions (in code)
    ↓
PermissionHelper.cs
    - User.Read
    - Mail.Read
    - Mail.Send
    - Calendars.Read
    - ... (16 total)

Active Permissions (in config)
    ↓
appsettings.json
    - User.Read (initially)
    - ... (added dynamically)
```

## Files Modified

### 1. `src/Office365MCP/appsettings.json`
- **Before:** 14 permissions pre-configured
- **After:** Only `User.Read` configured
- **Impact:** Simpler default, cleaner config

### 2. `src/Office365MCP/Auth/GraphAuthService.cs`
- **Added:** Smart default for missing scopes
- **Added:** Ensures `User.Read` always included
- **Impact:** More robust, handles empty config gracefully

### 3. `tests/Office365MCP.Tests/Auth/GraphAuthServiceTests.cs`
- **Changed:** Test for missing scopes now verifies default behavior
- **Was:** Expected exception when scopes missing
- **Now:** Expects successful creation with `User.Read` default

### 4. `SETUP.md`
- **Updated:** Configuration examples show minimal setup
- **Updated:** Dynamic permissions listed as Option 1 (recommended)
- **Updated:** Emphasizes that permissions are added as needed

## Benefits

### For New Users
- ✅ **Minimal setup** - Just start using tools
- ✅ **No confusion** - Don't need to know which permissions to configure
- ✅ **Learn by doing** - Discover permissions as you need them

### For Existing Users
- ✅ **Clean migration** - Can reset to minimal config
- ✅ **Better understanding** - See exactly which permissions you actually use
- ✅ **Flexibility** - Can still pre-configure if desired

### For Developers
- ✅ **Single source of truth** - Permissions defined in code
- ✅ **No duplication** - Don't maintain permissions in multiple places
- ✅ **Easier maintenance** - Add new permissions in one place

## Migration Guide

If you're upgrading from a previous configuration:

### Option 1: Keep Your Existing Config
Your existing `appsettings.json` with multiple scopes will continue to work perfectly. The server will use whatever scopes you've configured.

### Option 2: Reset to Minimal Config
1. Edit `appsettings.json`
2. Replace the `Scopes` array with just `["User.Read"]`
3. Restart the server
4. Permissions will be added dynamically as you use tools

### Option 3: Complete Fresh Start
1. Delete `%LocalAppData%\Office365MCP\office365mcp.cache` (token cache)
2. Reset `appsettings.json` to minimal config
3. Restart server
4. Re-authenticate with just User.Read
5. Grant additional permissions as needed

## Permission Discovery Flow

```
User → Claude → MCP Server → Microsoft Graph API
                    ↓              ↓
                    ↓         403 Forbidden
                    ↓              ↓
                    ↓    "Missing: Mail.Read"
                    ↓              ↓
           PermissionHelper detects error
                    ↓
         "Add Mail.Read permission? (yes/no)"
                    ↓
              User: yes
                    ↓
      appsettings.json updated:
      {
        "Scopes": [
          "User.Read",
          "Mail.Read"  ← Added
        ]
      }
                    ↓
         Token cache cleared
                    ↓
      Browser opens for consent
                    ↓
         User grants Mail.Read
                    ↓
        New token acquired
                    ↓
      Original request retries
                    ↓
            Success!
```

## Verification

- ✅ Build: Success (0 warnings, 0 errors)
- ✅ Tests: 15/15 Passing
- ✅ Default config: Only User.Read
- ✅ Smart defaults: Empty scopes → User.Read
- ✅ Dynamic add: Works as expected

## Summary

The Office 365 MCP Server now has a **minimal initial configuration** that starts with just the essential `User.Read` permission. All other permissions are:

1. **Defined in code** - Single source of truth in `PermissionHelper.cs`
2. **Added dynamically** - Prompted when needed, not pre-configured
3. **Stored in config** - Saved to `appsettings.json` after user consent

This provides a **better user experience** with less upfront configuration and clearer permission management.
