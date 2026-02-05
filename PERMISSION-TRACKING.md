# Permission Tracking & Reversion

## Overview

The Office 365 MCP Server now **automatically tracks all permission changes** made during runtime. You can view the history of changes and revert any dynamically added permissions back to the default set.

## Features

### ✅ Automatic Tracking
- Every dynamically added permission is logged to a tracking file
- Includes timestamp, reason, and context for each change
- Stored persistently at `%LocalAppData%\Office365MCP\permission-changes.json`

### ✅ View History
- See all permissions that were added dynamically
- View when and why each permission was requested
- Distinguish between active and reverted permissions

### ✅ Revert Changes
- Revert specific permissions individually
- Revert all dynamic permissions at once
- Default permissions (User.Read, Mail.Read, etc.) cannot be reverted

### ✅ MCP Tools
- Five new tools available through Claude for managing permissions
- No need to manually edit config files or tracking data
- Safe operations with validation and error handling

## Available MCP Tools

### 1. ListCurrentPermissions
**Purpose:** View all currently active permissions

```
User: "What permissions does this MCP server have?"
Claude: [Calls ListCurrentPermissions]

Response:
Current Permissions:

Default Permissions (cannot be reverted):
  ✓ User.Read
  ✓ Mail.Read
  ✓ Calendars.Read
  ✓ Chat.Read
  ✓ Notes.Read

Dynamically Added Permissions (can be reverted):
  + Mail.Send
  + Calendars.ReadWrite

Total: 7 permission(s) active
```

### 2. ViewPermissionHistory
**Purpose:** See the complete history of permission changes

```
User: "Show me the history of permission changes"
Claude: [Calls ViewPermissionHistory]

Response:
Permission Change History (3 change(s)):
Last Modified: 2026-02-05 09:30:15 UTC

Active Changes:
  [2026-02-05 09:25:10] + Mail.Send
    Reason: Added dynamically - Send emails

  [2026-02-05 09:28:42] + Calendars.ReadWrite
    Reason: Added dynamically - Read and create calendar events

Reverted Changes:
  [2026-02-05 09:30:15] - Team.ReadBasic.All
    Reason: Reverted by user
```

### 3. GetDefaultPermissions
**Purpose:** View the baseline permissions that are always active

```
User: "What are the default permissions?"
Claude: [Calls GetDefaultPermissions]

Response:
Default Permissions (always active, cannot be reverted):

  ✓ User.Read
  ✓ Mail.Read
  ✓ Calendars.Read
  ✓ Chat.Read
  ✓ Notes.Read

These permissions provide basic read and search capabilities across:
  - Mail (read and search emails)
  - Calendar (view events)
  - Teams (read chats and search messages)
  - OneNote (read notebooks)
```

### 4. RevertPermission
**Purpose:** Remove a specific dynamically added permission

```
User: "Remove the Mail.Send permission"
Claude: [Calls RevertPermission with permission="Mail.Send"]

Response:
{
  "success": true,
  "message": "Permission 'Mail.Send' has been reverted. Token cache cleared. The permission will no longer be included in future authentication requests."
}
```

### 5. RevertAllDynamicPermissions
**Purpose:** Reset to default permissions only

```
User: "Reset all permissions to defaults"
Claude: [Calls RevertAllDynamicPermissions]

Response:
{
  "success": true,
  "message": "All dynamic permissions have been reverted. Server reset to default permissions only.",
  "remainingPermissions": [
    "User.Read",
    "Mail.Read",
    "Calendars.Read",
    "Chat.Read",
    "Notes.Read"
  ]
}
```

## Tracking File Format

Location: `%LocalAppData%\Office365MCP\permission-changes.json`

```json
{
  "Changes": [
    {
      "Permission": "Mail.Send",
      "Timestamp": "2026-02-05T09:25:10.1234567Z",
      "Reason": "Added dynamically - Send emails",
      "AddedBy": "MCP Server",
      "IsReverted": false
    },
    {
      "Permission": "Team.ReadBasic.All",
      "Timestamp": "2026-02-05T09:30:15.7654321Z",
      "Reason": "Reverted by user",
      "AddedBy": "MCP Server",
      "IsReverted": true
    }
  ],
  "LastModified": "2026-02-05T09:30:15.7654321Z"
}
```

### Fields Explained

- **Permission**: The Graph API scope name
- **Timestamp**: When the change occurred (UTC)
- **Reason**: Why the permission was added or reverted
- **AddedBy**: Always "MCP Server" (for audit trail)
- **IsReverted**: `false` for active permissions, `true` for reverted ones

## Usage Scenarios

### Scenario 1: Track What Was Added

```
# Day 1: User adds Mail.Send while using the server
User: "Send an email to john@example.com"
Server: Prompts for Mail.Send permission
User: Grants permission
Server: Tracks the change

# Day 7: User wants to audit changes
User: "What permissions have been added since I started?"
Claude: [Calls ViewPermissionHistory]
User: Sees Mail.Send was added on Day 1
```

### Scenario 2: Revoke Specific Permission

```
User: "I don't want the MCP server to send emails anymore"
Claude: [Calls RevertPermission with "Mail.Send"]
Server: Removes Mail.Send from config
Server: Clears token cache
Server: Tracks the reversion
Next auth: Token acquired without Mail.Send
```

### Scenario 3: Fresh Start

```
User: "I've granted too many permissions, start over"
Claude: [Calls RevertAllDynamicPermissions]
Server: Removes all dynamic permissions
Server: Keeps only the 5 default read permissions
Server: Clears token cache
Server: All changes tracked as reverted
Next startup: Server has baseline permissions only
```

### Scenario 4: Audit Before Important Operation

```
# Before a critical meeting
User: "What permissions does this MCP server currently have?"
Claude: [Calls ListCurrentPermissions]
User: Reviews the list
User: "Remove the Team.ReadBasic.All permission"
Claude: [Calls RevertPermission]
User: Confident that server has minimal necessary permissions
```

## How It Works

### When a Permission is Added

```
1. Tool operation fails with 403 Forbidden
2. PermissionHelper detects missing permission
3. User is prompted to add the permission
4. User consents → permission added to config
5. Change is tracked to permission-changes.json:
   {
     "Permission": "Mail.Send",
     "Timestamp": "2026-02-05T09:25:10Z",
     "Reason": "Added dynamically - Send emails",
     "IsReverted": false
   }
6. Token cache is cleared
7. New token acquired with additional permission
```

### When a Permission is Reverted

```
1. User calls RevertPermission tool
2. PermissionHelper validates it's not a default permission
3. Permission is removed from appsettings.json
4. Change is tracked to permission-changes.json:
   {
     "Permission": "Mail.Send",
     "Timestamp": "2026-02-05T15:30:00Z",
     "Reason": "Reverted by user",
     "IsReverted": true
   }
5. Token cache is cleared
6. Next token acquisition excludes the reverted permission
```

## Default Permissions (Cannot Revert)

These 5 permissions are **always included** and cannot be reverted:

| Permission | Purpose | Why Required |
|------------|---------|--------------|
| `User.Read` | Read user profile | Required for Graph API access |
| `Mail.Read` | Read emails | Core mail functionality |
| `Calendars.Read` | View calendar | Core calendar functionality |
| `Chat.Read` | Read Teams chats | Core Teams functionality |
| `Notes.Read` | Read OneNote | Core OneNote functionality |

**Attempting to revert a default permission:**
```
User: "Revert User.Read"
Claude: [Calls RevertPermission with "User.Read"]
Response: ✗ Cannot revert 'User.Read' - it's a default permission
```

## Security & Privacy

### Audit Trail
- Complete history of all permission changes
- Timestamps for accountability
- Reasons for each change documented

### User Control
- User must approve all permission additions
- User can revoke permissions at any time
- Transparent - history is viewable

### Data Location
- Tracking file: `%LocalAppData%\Office365MCP\permission-changes.json`
- Same directory as token cache (secure, user-specific)
- Not synced or shared

### Safe Operations
- Default permissions protected from deletion
- Config file atomically updated
- Token cache cleared after each change
- No risk of corrupted state

## Technical Details

### File Operations
- **Read**: Synchronous, on-demand
- **Write**: Asynchronous, after every change
- **Format**: JSON with pretty-printing
- **Atomicity**: Entire file rewritten each time

### Token Cache Behavior
- Cleared after adding permission
- Cleared after reverting permission
- Forces re-authentication
- New token reflects current permission set

### Error Handling
- Failed tracking doesn't block permission changes
- Warnings logged but operations continue
- Tracking file creation is automatic
- Missing file treated as empty history

## Benefits

### For Users
- ✅ **Transparency** - Know exactly what permissions you've granted
- ✅ **Control** - Revoke permissions anytime
- ✅ **Audit** - Review history of changes
- ✅ **Privacy** - Minimize granted permissions

### For Administrators
- ✅ **Accountability** - Track who granted what and when
- ✅ **Compliance** - Audit trail for security reviews
- ✅ **Troubleshooting** - Understand permission-related issues
- ✅ **Governance** - Monitor permission sprawl

### For Developers
- ✅ **Debugging** - See what permissions are active
- ✅ **Testing** - Easy to reset to baseline
- ✅ **Maintenance** - Clear history of changes
- ✅ **Support** - Help users diagnose issues

## Common Questions

**Q: Can I manually edit the tracking file?**
A: Yes, but it's not recommended. Use the MCP tools instead. The tracking file is for reference, not configuration.

**Q: What happens if I delete the tracking file?**
A: The server will work fine. A new tracking file will be created on the next permission change. Historical data is lost but no functional impact.

**Q: Can I revert a permission and then add it again?**
A: Yes. Reverting removes it, but if you use a tool that needs it, you'll be prompted to add it again. Both changes are tracked.

**Q: Do reverted permissions stay in the history?**
A: Yes. The history shows both additions and reversions. This provides a complete audit trail.

**Q: What if I want to permanently hide a permission from history?**
A: You can manually edit the tracking JSON file to remove specific entries, but this breaks the audit trail.

**Q: How do I know if a permission needs admin consent?**
A: When prompted to add a permission, you'll see "⚠ WARNING: This permission requires administrator consent" if applicable.

## Summary

The permission tracking feature provides **complete transparency and control** over the permissions granted to the Office 365 MCP Server. You can:

- 📊 **View** current permissions and historical changes
- ➕ **Add** permissions dynamically when needed
- ➖ **Revert** specific permissions or reset to defaults
- 📝 **Audit** the complete history of permission changes
- 🔒 **Protect** default permissions from accidental removal

All through simple MCP tools accessible via Claude - no manual file editing required!
