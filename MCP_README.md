# Outlook MCP Server

This directory contains a complete Model Context Protocol (MCP) server for Microsoft Outlook integration.

## Quick Start

1. Set up Azure app registration with required permissions
2. Copy `.env.example` to `.env` and configure your credentials
3. Install dependencies: `npm install`
4. Build the project: `npm run build`
5. Start the server: `npm start`

## Features

- Email management (read, search, send)
- Calendar integration (view, create events)
- Profile and folder access
- Full MCP compliance with tools and resources

## Testing

Run `npm test` to verify the server starts correctly.

## Integration

Use the provided `mcp-config.json` for easy integration with MCP clients.