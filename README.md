# Office 365 MCP Server

A Model Context Protocol (MCP) server implementation in .NET C# that provides tools for interacting with Microsoft 365 services including Outlook (mail + calendar), Teams, and OneNote through the Microsoft Graph API.

## Features

- ✉️ **Outlook Mail** - Read, search, and send emails
- 📅 **Outlook Calendar** - View and create calendar events
- 💬 **Microsoft Teams** - Access teams, channels, chats, and messages
- 📓 **OneNote** - Read and create OneNote notebooks, sections, and pages
- 🔐 **Secure Authentication** - Interactive browser-based OAuth2 with token caching via Windows DPAPI
- 🚀 **No App Registration** - Uses Microsoft Graph Command Line Tools well-known Client ID
- ⚡ **Dynamic Permissions** - Automatically prompts to add missing permissions when needed

## Quick Start

### Prerequisites

- .NET 10.0 SDK or later
- Windows OS (for DPAPI token caching)
- Microsoft 365 account

### Build

```bash
dotnet build
```

### Test

```bash
dotnet test
```

### Configure Permissions

You have two options for managing permissions:

#### Option 1: Dynamic Permissions (Recommended for Getting Started)

Just start using the tools! When you try to use a tool that requires a permission you haven't granted:

1. The server detects the missing permission
2. You're prompted: "Would you like to add this permission?"
3. Type 'yes' and a browser opens for consent
4. The operation retries automatically

**Example:**
```
User: "List my Teams channels"
Server: Permission Required - Team.ReadBasic.All
Server: Would you like to add this permission? (yes/no)
User: yes
Server: ✓ Permission added, opening browser for consent...
```

See [DYNAMIC-PERMISSIONS.md](DYNAMIC-PERMISSIONS.md) for details.

#### Option 2: Batch Configuration

Use the interactive PowerShell script to configure all permissions upfront:

```powershell
.\setup-permissions.ps1
```

Or manually edit `src/Office365MCP/appsettings.json` to add/remove scopes.

### Integration with Claude

#### Claude Desktop

Add to your `claude_desktop_config.json`:

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

#### Claude Code

Add to your `claude_code_config.json`:

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
- `ListEmails` - List recent emails with optional count
- `GetEmail` - Get a specific email by ID
- `SearchEmails` - Search emails using KQL (Keyword Query Language)
- `SendEmail` - Send an email with subject and body
- `ListMailFolders` - List all mail folders

### Outlook Calendar
- `ListCalendarEvents` - List events within a date range
- `GetCalendarEvent` - Get a specific event by ID
- `CreateCalendarEvent` - Create a new calendar event with attendees
- `ListCalendars` - List all calendars

### Microsoft Teams
- `ListJoinedTeams` - List all teams the user has joined
- `ListChannels` - List channels in a specific team
- `ListChannelMessages` - List recent messages in a channel
- `SearchTeamsMessages` - Search messages across all Teams/channels/chats
- `SendChannelMessage` - Post a message to a channel
- `ListChats` - List user's chats
- `ListChatMessages` - List messages in a chat
- `SendChatMessage` - Send a message to a chat

### OneNote
- `ListNotebooks` - List all OneNote notebooks
- `ListSections` - List sections in a notebook
- `ListPages` - List pages in a section
- `GetPageContent` - Get the HTML content of a page
- `CreatePage` - Create a new page in a section

## Architecture

```
src/Office365MCP/
├── Program.cs              # MCP server startup and DI configuration
├── appsettings.json        # Configuration (Client ID, scopes)
├── Auth/
│   ├── IGraphAuthService.cs
│   └── GraphAuthService.cs # MSAL authentication with token caching
├── Graph/
│   ├── IGraphClientService.cs
│   └── GraphClientService.cs # HTTP client for Microsoft Graph API
└── Tools/
    ├── OutlookMailTools.cs
    ├── OutlookCalendarTools.cs
    ├── TeamsTools.cs
    └── OneNoteTools.cs
```

## How It Works

1. **Authentication**: Uses MSAL.NET (Microsoft Authentication Library) with the well-known Graph Command Line Tools Client ID
2. **Token Caching**: Securely stores tokens in Windows Credential Store using DPAPI (Data Protection API)
3. **First Run**: Opens browser for interactive sign-in and consent
4. **Subsequent Runs**: Uses cached tokens silently until expiration
5. **401 Retry**: Automatically clears cache and re-authenticates on 401 errors
6. **MCP Protocol**: Exposes tools via stdio transport for MCP clients

## Permissions

The server uses delegated permissions (acts on behalf of the signed-in user). Some permissions require admin consent:

- ✅ **User-consentable**: Mail, Calendar, Chat, OneNote
- ⚠️ **Admin consent required**: Teams channel reading and posting

See [SETUP.md](SETUP.md) for full permission details.

## Documentation

- [SETUP.md](SETUP.md) - Detailed setup and troubleshooting guide
- [CLAUDE.md](CLAUDE.md) - Project coding standards and conventions

## Security

- Tokens stored encrypted via Windows DPAPI
- No client secrets required (public client flow)
- User-specific token cache
- Microsoft's official well-known Client ID

## Technology Stack

- **.NET 10.0** - Application framework
- **ModelContextProtocol** - MCP SDK
- **Microsoft.Identity.Client** - MSAL for OAuth2
- **Microsoft.Graph** - Microsoft Graph API
- **xUnit** - Testing framework
- **NSubstitute** - Mocking library

## License

This project is provided as-is for educational and demonstration purposes.

## Contributing

Contributions are welcome! Please ensure:
- Code follows the standards in [CLAUDE.md](CLAUDE.md)
- All tests pass (`dotnet test`)
- New features include tests
- File-scoped namespaces are used
- Standard .NET naming conventions are followed

## Support

For issues and questions:
- Review [SETUP.md](SETUP.md) for troubleshooting
- Check [Microsoft Graph API documentation](https://docs.microsoft.com/en-us/graph/)
- Review [Model Context Protocol docs](https://modelcontextprotocol.io/)
