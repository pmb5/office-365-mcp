# Dynamic Permission Requesting

The Office 365 MCP Server includes automatic permission detection and requesting. When you try to use a tool that requires a permission you haven't granted yet, the server will detect the error and prompt you to add the missing permission.

## How It Works

1. **You use a tool** - For example, you ask Claude to list your Teams channels
2. **Permission error detected** - The server detects that you don't have `Team.ReadBasic.All` permission
3. **Interactive prompt** - You're asked if you want to add the permission:

```
╔═══════════════════════════════════════════════════════════════╗
║                  PERMISSION REQUIRED                          ║
╚═══════════════════════════════════════════════════════════════╝

  The operation you requested requires the following permission:

    Permission: Team.ReadBasic.All
    Purpose:    Read basic team information

  Would you like to add this permission to your configuration?
  (A browser window will open to grant consent)

  Type 'yes' to add permission, or 'no' to cancel:
  >
```

4. **Permission added** - If you type 'yes':
   - The permission is added to `appsettings.json`
   - Your token cache is cleared
   - A browser window opens for you to consent
   - The operation retries automatically

5. **Operation completes** - Your original request succeeds with the new permission

## Example Scenarios

### Scenario 1: Reading Teams Channels

```
User: "List my Teams channels"
Claude: [Calls ListChannels tool]
Server: Permission error detected - Team.ReadBasic.All required
Server: [Prompts user]
User: yes
Server: ✓ Permission added, browser opens
User: [Grants consent in browser]
Server: [Retries operation]
Claude: "Here are your Teams channels..."
```

### Scenario 2: Sending Email

```
User: "Send an email to john@example.com"
Claude: [Calls SendEmail tool]
Server: Permission error detected - Mail.Send required
Server: [Prompts user]
User: yes
Server: ✓ Permission added, opens consent page
User: [Grants consent]
Server: [Retries send]
Claude: "Email sent successfully"
```

## Supported Permissions

The following permissions can be automatically requested:

### User Profile
- `User.Read` - Read user profile

### Outlook Mail
- `Mail.Read` - Read email messages
- `Mail.Send` - Send emails
- `Mail.ReadBasic` - Read basic email properties
- `Mail.ReadWrite` - Read and write email

### Outlook Calendar
- `Calendars.Read` - Read calendar events
- `Calendars.ReadWrite` - Read and create calendar events

### Microsoft Teams
- `Team.ReadBasic.All` - Read basic team information (requires admin consent)
- `Channel.ReadBasic.All` - Read basic channel information (requires admin consent)
- `ChannelMessage.Read.All` - Read channel messages (requires admin consent)
- `ChannelMessage.Send` - Send channel messages (requires admin consent)
- `Chat.Read` - Read chat messages
- `Chat.ReadWrite` - Read and send chat messages

### OneNote
- `Notes.Read` - Read OneNote notebooks
- `Notes.ReadWrite` - Read and write OneNote
- `Notes.Create` - Create OneNote pages

## Admin Consent Requirements

Some permissions (marked above) require administrator consent. When you try to add these:

1. The server will warn you: "⚠ WARNING: This permission requires administrator consent"
2. You can still add it to your configuration
3. When the browser opens, you'll need an admin to grant the permission
4. If you're not an admin, share the consent URL with your IT admin

## Configuration File

Permissions are stored in `src/Office365MCP/appsettings.json`:

```json
{
  "Graph": {
    "ClientId": "14d82eec-204b-4c2f-b7e8-296a70dab67e",
    "TenantId": "organizations",
    "Scopes": [
      "User.Read",
      "Mail.Read",
      "Team.ReadBasic.All"  // Added dynamically when you used a Teams tool
    ]
  }
}
```

## Benefits

### No Upfront Configuration
- Don't need to run the PowerShell script
- Don't need to know which permissions you'll need
- Just start using tools naturally

### Just-In-Time Permissions
- Only request permissions when actually needed
- Principle of least privilege
- Clear context for why each permission is needed

### User Control
- You approve each permission
- Can decline if you don't want to grant access
- Permission descriptions explain what they do

## Comparing Approaches

### Approach 1: PowerShell Script (setup-permissions.ps1)
**When to use:** Configure all permissions upfront
- ✅ One-time setup
- ✅ No interruptions during use
- ✅ Good for known workflows
- ❌ Need to know which permissions ahead of time
- ❌ Grants more permissions than immediately needed

### Approach 2: Dynamic Permission Requesting (Automatic)
**When to use:** Discover permissions as you go
- ✅ No upfront configuration
- ✅ Only grant what you actually use
- ✅ Clear context for each permission
- ❌ Interrupts workflow when new permission needed
- ❌ Requires re-authentication for each new permission

### Recommended: Use Both
1. Start with no configuration (or just `User.Read`)
2. Use tools naturally - add permissions as needed
3. Once you know your workflow, run `setup-permissions.ps1` to add remaining permissions in bulk

## Technical Details

### Error Detection

The server detects permission errors by:
1. Catching 403 Forbidden HTTP responses from Microsoft Graph API
2. Checking the error message for permission-related keywords
3. Extracting the specific permission name from the error

### Permission Extraction

Error messages are parsed to find permission names:
- Checks JSON error responses
- Looks for permission names in error text
- Only prompts for permissions in the supported list

### Configuration Update

When you grant a permission:
1. `appsettings.json` is read and parsed
2. The new scope is added to the `Graph:Scopes` array
3. File is written back with proper JSON formatting
4. Token cache is cleared to force re-authentication

### Automatic Retry

After adding a permission:
1. Token cache is cleared
2. Next auth request triggers interactive browser login
3. You consent to the new permission
4. Token is cached with new scopes
5. Original operation retries automatically
6. No need to repeat your request to Claude

## Troubleshooting

### Permission not detected
If the server doesn't prompt for a missing permission:
- Check the error message in logs
- The permission might not be in the supported list
- Manually add it using `setup-permissions.ps1`

### Browser doesn't open
- The browser should open automatically
- If not, check the consent URL in the error message
- Copy and paste it into your browser manually

### Admin consent fails
- Contact your IT administrator
- Share the consent URL with them
- They need to grant the permission for your organization

### Permission already exists
- The server checks if the permission is already configured
- Won't prompt if it's already in `appsettings.json`
- Check the file to verify the scope is present

## Security Considerations

- Each permission requires explicit user approval
- Permissions are only added to local configuration
- Token cache uses DPAPI encryption (Windows)
- No permissions are granted without your knowledge
- All permissions use delegated access (on behalf of you)

## See Also

- [SETUP.md](SETUP.md) - Manual permission configuration
- [README.md](README.md) - General overview
- [setup-permissions.ps1](setup-permissions.ps1) - Batch permission configuration script
