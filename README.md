# Simply Outlook MCP Server

A Model Context Protocol (MCP) server that enables AI assistants to interact with Microsoft Outlook calendars and emails through the Microsoft Graph API.

## 🚀 Features

### 📅 Calendar Management
- **Get Calendar Events** - Retrieve calendar events within specified date ranges
- **Create Calendar Events** - Create personal calendar events
- **Create Events with Invites** - Create calendar events and send invitations to attendees
- **Update Calendar Events** - Modify existing calendar events (subject, time, location, etc.)

### 📧 Email Operations
- **Get Messages** - Fetch recent Outlook messages with date filtering
- **Get Message Content** - Retrieve full content of specific messages by ID
- **Search Messages** - Search emails by keywords (sender, subject, content)
- **Send Messages** - Send new emails to specified recipients
- **Reply to Messages** - Reply to existing email messages
- **Pagination Support** - Handle large result sets with skip/limit parameters

### ⚙️ Configuration
- **Selective Tool Control** - Enable/disable individual tools via environment variables
- **Environment Configuration** - Flexible setup via environment variables
- **TypeScript Support** - Full type definitions included

## 📋 Prerequisites

- [**Node.js**](https://nodejs.org/) 20.19.0 or higher
- **Azure AD Application**:
  - `Calendars.ReadWrite` - Read and write calendar events
  - `Mail.Read` - Read email messages
  - `Mail.Send` - Send email messages
  - `User.Read` - Read user profile information

## 🛠️ Installation

### Using npx (Recommended)
No installation required! Simply use `npx` to run the latest version:

### Global Installation (Optional)
If you prefer to install globally:

```bash
npm install -g simply-outlook-mcp
```

### From Source
For development or customization:

```bash
npm install
npm run build
```

## 🔧 Setup

> [!NOTE]
> **Security Best Practice**: We recommend creating your own Azure AD application rather than using a client ID from an unknown 3rd party or untrusted publisher.
>
> This MCP server runs locally. The code that uses the Microsoft Graph access tokens stays on your machine. The Azure AD / Entra "client ID" you supply only tells the directory *which* app registration to issue tokens for. If you reuse an arbitrary or third‑party client ID, you do **not** automatically leak your tokens to that party, but you *are* accepting several avoidable risks tied to the *app registration* that they control.
>
> **What a third party CAN change later if you use their client ID:**
> - Display name / branding you see during future consent prompts (could lower your guard)
> - The set of requested delegated scopes (you may click "Accept" again out of habit)
> - Publisher domain / verification state (influences user trust signals)
> - Authentication settings (enabling flows they host to phish you for tokens)
>

### 1. Azure AD Application Registration

1. Go to the [Azure Portal](https://portal.azure.com/)
2. Navigate to **Microsoft Entra ID** > **Manage** > **App registrations**
3. Click **New registration**
4. Configure:
   - **Name**: (Any name you like, will show up in the consent dialog)
   - **Supported account types**: `Accounts in any organizational directory (Any Microsoft Entra ID tenant - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)`
   - **Redirect URI**: Leave blank
5. After creation, in the application resource view, note down client ID in the **Overview** page:
   - **Application (client) ID**
6. In the application resource view, go to **Manage** > **Authentication** page:
   - Toggle `Enable the following mobile and desktop flows` to `Yes`

> [!NOTE]
> The scopes currently requested by this MCP are designed to work with personal Microsoft account (MSA).

### 2. Authentication Setup

```bash
npx simply-outlook-mcp --auth --client_id YOUR_CLIENT_ID
```

This will open a browser for device authentication. Sign in with your Microsoft account that has access to the Outlook data you want to manage.

## 🚀 Usage

### Using with MCP Clients

The server implements the Model Context Protocol and can be used with any MCP-compatible AI assistant.

### Example configuration for Claude Desktop:

```json
{
  "mcpServers": {
    "simply-outlook-mcp": {
      "command": "npx",
      "args": [
        "-y",
        "simply-outlook-mcp"
      ],
      "env": {
        "SIMPLY_OUTLOOK_MCP_CLIENT_ID": "your-client-id",
      }
    }
  }
}
```

### VS Code

```json
{
  "servers": {
    "simply-outlook-mcp": {
      "type": "stdio",
      "command": "npx",
      "args": [
        "-y",
        "simply-outlook-mcp"
      ],
      "env": {
        "SIMPLY_OUTLOOK_MCP_CLIENT_ID": "your-client-id"
      }
    }
  }
}
```

## Environment Variables

The following environment variables can be used to configure the MCP server:

### Required Variables

| Variable | Description | Example |
|----------|-------------|---------|
| `SIMPLY_OUTLOOK_MCP_CLIENT_ID` | Azure AD Application (client) ID | `12345678-1234-1234-1234-123456789012` |

### Optional Variables

| Variable | Description | Default | Example |
|----------|-------------|---------|---------|
| `SIMPLY_OUTLOOK_MCP_TENANT_ID` | Ignore this, for advance usage if you create single tenant or MSA only app | `common` | |
| `SIMPLY_OUTLOOK_MCP_DISABLED_TOOLS` | Comma-separated list of tools to disable | None | `get-calendar-events,send-outlook-message` |

### Tool Names for Disabling

You can disable specific tools by adding their names to the `SIMPLY_OUTLOOK_MCP_DISABLED_TOOLS` environment variable:

**Calendar Tools:**
- `get-calendar-events` - Retrieve calendar events
- `create-calendar-event` - Create personal calendar events  
- `create-calendar-event-with-invite` - Create events with invitations
- `update-calendar-event` - Modify existing calendar events

**Email Tools:**
- `get-outlook-messages` - Fetch recent messages
- `get-outlook-message-content` - Get full message content
- `search-outlook-messages` - Search emails by keywords
- `send-outlook-message` - Send new emails
- `reply-outlook-message` - Reply to existing emails

### Example Configuration

```json
{
  "mcpServers": {
    "outlook": {
      "command": "npx",
      "args": [
        "-y",
        "simply-outlook-mcp"
      ],
      "env": {
        "SIMPLY_OUTLOOK_MCP_CLIENT_ID": "12345678-1234-1234-1234-123456789012",
        "SIMPLY_OUTLOOK_MCP_DISABLED_TOOLS": "create-calendar-event-with-invite,update-calendar-event,send-outlook-message,reply-outlook-message"
      }
    }
  }
}
```

> [!TIP]
> **Disabling Tools**: You can disable tools you don't need for security or functionality reasons. For example, disable email sending tools (`send-outlook-message,reply-outlook-message`) if you only want read-only email access.

## License

MIT License - see [LICENSE](LICENSE) file for details.

## Support

- GitHub Issues: [Report bugs or request features](https://github.com/hmmroger/simply-outlook-mcp/issues)
- Repository: [https://github.com/hmmroger/simply-outlook-mcp](https://github.com/hmmroger/simply-outlook-mcp)
