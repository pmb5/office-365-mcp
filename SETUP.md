# Office 365 MCP Server - Setup Guide

## Prerequisites

- .NET 10.0 SDK or later
- Windows OS (for DPAPI token caching)
- Microsoft 365 account

## Building the Project

```bash
dotnet build
```

## Configuration

The MCP server is configured via `appsettings.json`:

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

### Configuration Parameters

- **ClientId**: The Microsoft Graph Command Line Tools well-known Client ID (no app registration required)
- **TenantId**: Use `"organizations"` for multi-tenant support
- **Scopes**: Array of Microsoft Graph API permission scopes (starts with just `User.Read`, others added dynamically as needed)

## Permissions Guide

The following table shows all available scopes, their purpose, which tools require them, and whether admin consent is needed:

| Scope | Description | Required For | Admin Consent |
|-------|-------------|--------------|---------------|
| `User.Read` | Read user profile | All tools | No |
| `Mail.Read` | Read user's mail | `ListEmails`, `GetEmail`, `SearchEmails` | No |
| `Mail.Send` | Send mail as the user | `SendEmail` | No |
| `Mail.ReadBasic` | Read basic mail properties | `ListEmails` (limited) | No |
| `Mail.ReadWrite` | Read and write user's mail | Full mail management | No |
| `Calendars.Read` | Read user's calendars | `ListCalendarEvents`, `GetCalendarEvent`, `ListCalendars` | No |
| `Calendars.ReadWrite` | Read and write user's calendars | `CreateCalendarEvent` | No |
| `Team.ReadBasic.All` | Read basic team info | `ListJoinedTeams` | Yes |
| `Channel.ReadBasic.All` | Read basic channel info | `ListChannels` | Yes |
| `ChannelMessage.Read.All` | Read channel messages | `ListChannelMessages`, `SearchChannelMessages` | Yes |
| `ChannelMessage.Send` | Send channel messages | `SendChannelMessage` | Yes |
| `Chat.Read` | Read user's chats | `ListChats`, `ListChatMessages`, `SearchChatMessages` | No |
| `Chat.ReadWrite` | Read and write user's chats | `SendChatMessage` | No |
| `Notes.Read` | Read user's OneNote notebooks | `ListNotebooks`, `ListSections`, `ListPages`, `GetPageContent` | No |
| `Notes.ReadWrite` | Read and write OneNote notebooks | All OneNote tools | No |
| `Notes.Create` | Create new OneNote pages | `CreatePage` | No |

### Configuring Permissions

You can configure permissions in three ways:

#### Option 1: Dynamic Permissions (Recommended)

**No configuration needed!** Just start using tools. When you try to use a tool that requires a permission you haven't granted:

1. The server detects the missing permission
2. Prompts you: "Would you like to add this permission?"
3. Type 'yes' and a browser opens for consent
4. The operation retries automatically

See [DYNAMIC-PERMISSIONS.md](DYNAMIC-PERMISSIONS.md) for details.

#### Option 2: Batch Configuration (PowerShell Script)

If you know which permissions you need upfront, use the interactive script:

```powershell
.\setup-permissions.ps1
```

The script will:
1. Display current permissions from `appsettings.json`
2. Let you toggle permissions on/off per service
3. Save updated scopes to `appsettings.json`
4. Generate admin consent URL if needed

#### Option 3: Manual Configuration

Edit `appsettings.json` directly and add/remove scopes in the `Graph:Scopes` array.

### Admin Consent

Some permissions (marked "Yes" in the table above) require admin consent. These are typically organization-wide permissions related to Teams.

**To obtain admin consent:**

1. Use the permissions script or manually generate a consent URL:
```
https://login.microsoftonline.com/organizations/v2.0/adminconsent?client_id=14d82eec-204b-4c2f-b7e8-296a70dab67e&scope=<space-separated-scopes>
```

2. Have your tenant administrator visit this URL and grant consent
3. After consent is granted, the MCP server will be able to use those permissions

## First Run

When you first run the MCP server, it will:

1. Open your default web browser for interactive authentication
2. Prompt you to sign in with your Microsoft 365 account
3. Request consent for the configured permissions
4. Cache the access token securely in the Windows credential store (DPAPI)

Subsequent runs will use the cached token silently until it expires.

## Running the Server

The MCP server is designed to be run as a subprocess by MCP clients (e.g., Claude Desktop, Claude Code).

```bash
dotnet run --project src/Office365MCP/Office365MCP.csproj
```

## Claude Desktop Integration

Add the following to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "office-365": {
      "command": "dotnet",
      "args": [
        "run",
        "--project",
        "D:\\code\\office-365-mcp\\src\\Office365MCP\\Office365MCP.csproj"
      ]
    }
  }
}
```

Or, if you've published the app:

```json
{
  "mcpServers": {
    "office-365": {
      "command": "D:\\code\\office-365-mcp\\src\\Office365MCP\\bin\\Debug\\net10.0\\Office365MCP.exe"
    }
  }
}
```

## Claude Code Integration

Add the following to your `claude_code_config.json`:

```json
{
  "mcpServers": {
    "office-365": {
      "command": "dotnet",
      "args": [
        "run",
        "--project",
        "D:\\code\\office-365-mcp\\src\\Office365MCP\\Office365MCP.csproj"
      ]
    }
  }
}
```

## Available Tools

### Outlook Mail
- `ListEmails` - List recent emails
- `GetEmail` - Get a specific email by ID
- `SearchEmails` - Search emails using KQL
- `SendEmail` - Send an email
- `ListMailFolders` - List mail folders

### Outlook Calendar
- `ListCalendarEvents` - List events in a date range
- `GetCalendarEvent` - Get a specific event by ID
- `CreateCalendarEvent` - Create a new calendar event
- `ListCalendars` - List all calendars

### Teams
- `ListJoinedTeams` - List user's teams
- `ListChannels` - List channels in a team
- `ListChannelMessages` - List messages in a channel
- `SearchChannelMessages` - Search messages in a channel
- `SearchTeamsMessages` - Search messages across all Teams
- `SendChannelMessage` - Post message to a channel
- `ListChats` - List user's chats
- `ListChatMessages` - List messages in a chat
- `SearchChatMessages` - Search messages in a chat
- `SendChatMessage` - Send a chat message

### OneNote
- `ListNotebooks` - List all notebooks
- `ListSections` - List sections in a notebook
- `ListPages` - List pages in a section
- `GetPageContent` - Get HTML content of a page
- `CreatePage` - Create a new page

## Troubleshooting

### Authentication Errors

**Error: "AADSTS65001: The user or administrator has not consented"**
- Solution: Run the app interactively to trigger consent flow, or use the admin consent URL

**Error: "AADSTS50076: Multi-factor authentication required"**
- Solution: Complete MFA in the browser window that opens

**Error: "Token has expired"**
- Solution: The MCP server automatically handles token refresh. If issues persist, clear the cache:

```csharp
// Token cache location
%LocalAppData%\Office365MCP\office365mcp.cache
```

### Clearing Token Cache

To force re-authentication, delete the token cache file:

```powershell
Remove-Item "$env:LOCALAPPDATA\Office365MCP\office365mcp.cache" -ErrorAction SilentlyContinue
```

### Permission Errors

**Error: "Insufficient privileges to complete the operation"**
- Solution: Ensure the required scope is in `appsettings.json` and you've consented to it
- For admin-required scopes, contact your tenant administrator

### Build Errors

**Error: Package 'ModelContextProtocol' not found**
- Solution: Enable prerelease packages:
```bash
dotnet add package ModelContextProtocol --prerelease
```

### Runtime Errors

**Error: "Graph:ClientId not configured"**
- Solution: Ensure `appsettings.json` exists and is copied to the output directory

**Error: "Failed to initialize token cache"**
- Solution: Ensure you're running on Windows (DPAPI support required)

## Security Notes

- Tokens are stored encrypted using Windows DPAPI (Data Protection API)
- Token cache is user-specific and cannot be accessed by other users
- The well-known Client ID (`14d82eec-204b-4c2f-b7e8-296a70dab67e`) is Microsoft's official Graph Command Line Tools app
- No client secrets are required (public client flow)

## Testing

Run the unit tests:

```bash
dotnet test
```

## Support

For issues and questions:
- Check the [Microsoft Graph API documentation](https://docs.microsoft.com/en-us/graph/)
- Review the [Model Context Protocol specification](https://modelcontextprotocol.io/)
- File issues in the project repository
