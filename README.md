# Outlook MCP Server

A Model Context Protocol (MCP) server implementation for Microsoft Outlook integration. Provides seamless access to Microsoft Graph API for email, calendar, and contact management.

## 🚀 Features

- **Email Management**: List, read, send, reply, forward, and delete emails
- **Calendar Management**: View, create, update, and delete events
- **Contact Management**: Manage contacts and contact groups
- **Authentication**: Secure OAuth2 with automatic token refresh
- **MCP Compatible**: Standard interface for AI tool integration

## 🛠️ Quick Setup

1. **Install dependencies**
   ```bash
   npm install
   ```

2. **Configure environment**
   ```bash
   cp .env.example .env
   # Edit .env with your Microsoft Graph API credentials
   ```

3. **Start the server**
   ```bash
   npm start
   ```

4. **Authenticate**
   Use the `outlook_auth` tool to complete OAuth2 authentication.

## 📋 Requirements

- Node.js 18+
- Microsoft Azure AD application with Graph API permissions
- Environment variables: `OUTLOOK_CLIENT_ID`, `OUTLOOK_CLIENT_SECRET`, `OUTLOOK_TENANT_ID`

## 🔧 Available Tools

| Tool | Description |
|------|-------------|
| `outlook_auth` | OAuth2 authentication |
| `outlook_list_emails` | List emails with filtering |
| `outlook_send_email` | Send emails |
| `outlook_list_calendar_events` | View calendar events |
| `outlook_create_event` | Create calendar events |
| `outlook_list_contacts` | Manage contacts |

## 📄 License

ISC

---

For detailed API documentation and setup instructions, see [NOTION_REPORT.md](NOTION_REPORT.md). 