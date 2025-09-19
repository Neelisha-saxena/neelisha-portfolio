# Neelisha Saxena - Portfolio & Outlook MCP Server

This repository contains Neelisha Saxena's portfolio website and an Outlook MCP (Model Context Protocol) server for Microsoft Outlook integration.

## Portfolio Website

A simple, elegant portfolio website showcasing Neelisha's experience as a Machine Learning Engineer at Microsoft.

## Outlook MCP Server

An MCP server that provides AI assistants with access to Microsoft Outlook functionality including emails, calendar events, and more.

### Features

- **Email Management**: Read, search, and send emails
- **Calendar Integration**: View and create calendar events
- **Profile Access**: Get user profile information
- **Folder Management**: Access mail folder structure

### Setup

1. **Azure App Registration**
   - Go to [Azure Portal](https://portal.azure.com)
   - Register a new application
   - Add required Microsoft Graph permissions:
     - `Mail.Read`
     - `Mail.Send`
     - `Calendars.Read`
     - `Calendars.ReadWrite`
     - `User.Read`

2. **Environment Configuration**
   ```bash
   cp .env.example .env
   # Edit .env with your Azure app credentials
   ```

3. **Install Dependencies**
   ```bash
   npm install
   ```

4. **Build and Run**
   ```bash
   npm run build
   npm start
   ```

### MCP Tools

#### Email Tools
- `get_emails`: Retrieve emails from inbox or other folders
- `send_email`: Send emails with optional CC/BCC

#### Calendar Tools
- `get_calendar_events`: Get calendar events within date range
- `create_calendar_event`: Create new calendar events

### MCP Resources
- `outlook://profile`: User profile information
- `outlook://folders`: Mail folder structure

### Usage with MCP Clients

The server communicates via stdio and can be integrated with MCP-compatible AI clients:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "node",
      "args": ["/path/to/outlook-mcp-server/dist/index.js"],
      "env": {
        "OUTLOOK_CLIENT_ID": "your_client_id",
        "OUTLOOK_CLIENT_SECRET": "your_client_secret",
        "OUTLOOK_TENANT_ID": "your_tenant_id"
      }
    }
  }
}
```

### Security Notes

- Store credentials securely in environment variables
- Use least-privilege permissions in Azure app registration
- Consider implementing token refresh mechanisms for long-running sessions

### Development

```bash
# Development mode with hot reload
npm run dev

# Type checking
npm run build
```
