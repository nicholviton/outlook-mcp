# Microsoft Outlook MCP Server

A Model Context Protocol (MCP) server that enables AI assistants to interact with Microsoft Outlook email and calendar through the Microsoft Graph API.

[![Download Latest Release](https://img.shields.io/github/v/release/XenoXilus/outlook-mcp?label=Download&color=blue)](https://github.com/XenoXilus/outlook-mcp/releases/latest)
[![Ko-fi](https://img.shields.io/badge/Ko--fi-Support-ff5f5f?logo=ko-fi)](https://ko-fi.com/xenoxilus)

## Features

- **Email Operations**: Read, search, send, reply to emails and download attachments
- **SharePoint Integration**: Access SharePoint files via sharing links or direct file IDs. Download files shared to you via emails. 
- **Calendar Management**: View and manage calendar events and appointments
- **Office Document Processing**: Parse PDF, Word, PowerPoint, and Excel files with extracted text content
- **Large File Support**: Automatic handling of files that exceed MCP response size limits

## Quick Start

**Choose your installation method:**

| Method | Best For |
|--------|----------|
| [DXT Extension](#installing-as-dxt-extension) | Claude Desktop users |
| [CLI Configuration](#using-with-cli-tools) | Claude Code, mcp CLI, other MCP clients |

> **Prerequisites**: Before installing, you'll need to [set up an Azure application](#azure-setup-guide) to get your Client ID and Tenant ID.

---

## Installation

### Installing as DXT Extension

For Claude Desktop users, DXT extensions provide the simplest installation experience.

**Option 1: Download Pre-built Extension**
1. Download `outlook-mcp.dxt` from the [Releases page](https://github.com/XenoXilus/outlook-mcp/releases)
2. In Claude Desktop, go to **Settings** → **Extensions**
3. Click **Install from file** and select the `.dxt` file
4. Enter your Azure Client ID, Tenant ID, and optional download directory when prompted

**Option 2: Build from Source**
1. Clone and install dependencies:
   ```bash
   git clone https://github.com/XenoXilus/outlook-mcp.git
   cd outlook-mcp
   npm install
   ```
2. Install the DXT CLI: `npm install -g @anthropic-ai/dxt`
3. Pack the extension:
   ```bash
   dxt pack . outlook-mcp.dxt
   ```
4. Install the generated `.dxt` file in Claude Desktop as above

---

### Using with CLI Tools

For CLI-based MCP clients (Claude Code, mcp CLI, etc.), configure the server directly.

**1. Clone and Install:**
```bash
git clone https://github.com/XenoXilus/outlook-mcp.git
cd outlook-mcp
npm install
```

**2. Configure your MCP client:**

Add the following to your MCP servers configuration (location varies by client):

```json
{
  "outlook-mcp": {
    "command": "node",
    "args": ["/absolute/path/to/outlook-mcp/server/index.js"],
    "env": {
      "AZURE_CLIENT_ID": "your-azure-client-id",
      "AZURE_TENANT_ID": "your-azure-tenant-id",
      "MCP_OUTLOOK_WORK_DIR": "/optional/download/directory"
    }
  }
}
```

**Common config file locations:**
- **Claude Code**: `~/.claude.json` or project-level `.mcp.json`
- **mcp CLI**: `~/.config/mcp/servers.json`

**3. Alternative: Use environment variables**

Instead of specifying `env` in the config, you can export the variables in your shell:

```bash
export AZURE_CLIENT_ID="your-azure-client-id"
export AZURE_TENANT_ID="your-azure-tenant-id"
export MCP_OUTLOOK_WORK_DIR="/optional/download/directory"
```

---

## Azure Setup Guide

To use this MCP server, you need to register an application in Microsoft Azure.

### For Business/Work Accounts (Recommended)

1. Go to the [Azure Portal](https://portal.azure.com/) and search for "App registrations".
2. Click **New registration**.
   - Name: `Outlook MCP` (or similar)
   - Supported account types: **Accounts in this organizational directory only** (Single tenant)
   - Redirect URI: Select **Web** and enter `http://localhost/callback`
3. Click **Register**.
4. Go to **Authentication** in the sidebar.
   - Under "Advanced settings", set **Allow public client flows** to **Yes**.
   - Click **Save**.
5. On the Overview page, copy:
   - **Application (client) ID** → This is your `AZURE_CLIENT_ID`
   - **Directory (tenant) ID** → This is your `AZURE_TENANT_ID`
6. Go to **API permissions** in the sidebar.
   - Click **Add a permission** -> **Microsoft Graph** -> **Delegated permissions**.
   - Add these permissions:
     - `Mail.Read`, `Mail.ReadWrite`, `Mail.Send`
     - `Calendars.Read`, `Calendars.ReadWrite`
     - `User.Read`, `MailboxSettings.Read`
     - `Files.Read.All`, `Files.ReadWrite.All`
     - `Sites.Read.All`, `Sites.ReadWrite.All`
     - `offline_access`
   - Click **Add permissions**.
   - (Optional) If you are an admin, click **Grant admin consent** to suppress consent prompts for users.

**Note:** No client secret is required (PKCE auth flow).

### For Personal Accounts (outlook.com, hotmail.com)

Personal Microsoft accounts can also register apps in Azure:

1. Sign in to the [Azure Portal](https://portal.azure.com/) with your personal Microsoft account (outlook.com, hotmail.com, etc.).
2. If prompted to create a directory, follow the steps to create a free Azure directory.
3. Follow the same steps as above for Business accounts.
4. When configuring, use **Accounts in any organizational directory and personal Microsoft accounts** for supported account types.

---

## Configuration Reference

### Environment Variables

| Variable | Required | Description |
|----------|----------|-------------|
| `AZURE_CLIENT_ID` | Yes | Your Azure AD application client ID |
| `AZURE_TENANT_ID` | Yes | Your Azure AD directory (tenant) ID |
| `MCP_OUTLOOK_WORK_DIR` | No | Directory for saving large files (defaults to system temp) |

### Large File Handling

When downloading large attachments or SharePoint files, the server automatically detects when the response would exceed the MCP 1MB limit and saves the content to local files instead.

- If `MCP_OUTLOOK_WORK_DIR` is set, large files are saved to this directory
- If not set, files are saved to the system temp directory
- Files are automatically named with timestamps to avoid conflicts
- Old files are periodically cleaned up to manage disk space

---

## Example Prompts

Once installed, you can ask the AI assistant things like:

**Email Management**
- "Show me my unread emails from this week"
- "Find all emails from John about the project proposal"
- "Send a reply to the last email from Sarah thanking her for the update"
- "Draft an email to the team summarizing today's meeting"

**Calendar**
- "What meetings do I have tomorrow?"
- "Schedule a 30-minute call with Alex next Tuesday afternoon"
- "Show me my availability for the rest of the week"

**Attachments & SharePoint**
- "Download and summarize the PDF attachment from the latest email from Finance"
- "Get the contents of this SharePoint link: [paste link]"
- "What files were attached to emails from Legal this month?"

**Office Document Processing**

The server automatically parses:
- **PDF files**: Extracts text content
- **Word documents** (.docx): Extracts text content
- **PowerPoint** (.pptx): Extracts slide text
- **Excel** (.xlsx): Parses data into structured format

---

## Authentication

The server uses OAuth 2.0 with PKCE for secure authentication:

1. First run will open a browser for Microsoft authentication
2. Tokens are encrypted and stored locally (uses OS keychain if available, otherwise encrypted file storage)
3. Automatic token refresh for long-term usage
4. No sensitive data stored in plain text

### Required Permissions

The app requests these Microsoft Graph permissions:

- `Mail.Read`, `Mail.ReadWrite`, `Mail.Send` - Email access
- `Calendars.Read`, `Calendars.ReadWrite` - Calendar access  
- `User.Read`, `MailboxSettings.Read` - User profile
- `Files.Read.All`, `Files.ReadWrite.All` - OneDrive/SharePoint files
- `Sites.Read.All`, `Sites.ReadWrite.All` - SharePoint sites
- `offline_access` - Refresh tokens

---

## Troubleshooting

### Large File Issues
- **Problem**: "Result exceeds maximum length" error
- **Solution**: Ensure `MCP_OUTLOOK_WORK_DIR` is set and writable
- **Alternative**: Files automatically save to system temp if work dir not configured

### Authentication Issues
- **Problem**: Authentication failures
- **Solution**: Verify Azure AD app permissions and client ID
- **Reset**: Clear stored tokens and re-authenticate

### SharePoint Access Issues
- **Problem**: Cannot access SharePoint files
- **Solution**: Ensure sharing links are valid and user has access permissions
- **Alternative**: Use direct file ID access if available

---

## Development

### Project Structure
```
outlook-mcp/
├── server/
│   ├── index.js              # Main MCP server
│   ├── auth/                 # Authentication management
│   ├── graph/                # Microsoft Graph API client
│   ├── schemas/              # MCP tool schemas
│   ├── tools/                # MCP tool implementations
│   │   ├── attachments/      # Attachment tools
│   │   ├── calendar/         # Calendar tools
│   │   ├── email/            # Email tools
│   │   ├── folders/          # Folder management
│   │   └── sharepoint/       # SharePoint tools
│   └── utils/                # Utility modules
└── package.json
```

### Running Tests
```bash
npm test                    # Run all tests
npm run test:watch          # Watch mode
npm run test:benchmark      # Performance benchmarks
```

### Debugging
```bash
npm run test:graph          # Test Graph API connection
```

---

## Support

If this tool saved you time, consider supporting the development!

[![Ko-fi](https://img.shields.io/badge/Ko--fi-Support-ff5f5f?logo=ko-fi)](https://ko-fi.com/xenoxilus)

---

## License

MIT License

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make changes with tests
4. Submit a pull request
