#!/usr/bin/env node

import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ErrorCode,
  ListResourcesRequestSchema,
  ListToolsRequestSchema,
  McpError,
  ReadResourceRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import { z } from "zod";
import { OutlookClient } from "./outlook-client.js";

// Configuration schema
const ConfigSchema = z.object({
  clientId: z.string(),
  clientSecret: z.string(),
  tenantId: z.string(),
  redirectUri: z.string().optional(),
});

class OutlookMCPServer {
  private server: Server;
  private outlookClient!: OutlookClient;

  constructor() {
    this.server = new Server(
      {
        name: "outlook-mcp-server",
        version: "1.0.0",
      },
      {
        capabilities: {
          tools: {},
          resources: {},
        },
      }
    );

    this.setupToolHandlers();
    this.setupResourceHandlers();
  }

  private setupToolHandlers() {
    this.server.setRequestHandler(ListToolsRequestSchema, async () => {
      return {
        tools: [
          {
            name: "get_emails",
            description: "Get emails from Outlook inbox",
            inputSchema: {
              type: "object",
              properties: {
                limit: {
                  type: "number",
                  description: "Number of emails to retrieve (default: 10)",
                  default: 10,
                },
                search: {
                  type: "string",
                  description: "Search query for emails",
                },
                folder: {
                  type: "string",
                  description: "Folder to search in (default: inbox)",
                  default: "inbox",
                },
              },
            },
          },
          {
            name: "send_email",
            description: "Send an email via Outlook",
            inputSchema: {
              type: "object",
              properties: {
                to: {
                  type: "array",
                  items: { type: "string" },
                  description: "Email recipients",
                },
                subject: {
                  type: "string",
                  description: "Email subject",
                },
                body: {
                  type: "string",
                  description: "Email body content",
                },
                cc: {
                  type: "array",
                  items: { type: "string" },
                  description: "CC recipients",
                },
                bcc: {
                  type: "array",
                  items: { type: "string" },
                  description: "BCC recipients",
                },
              },
              required: ["to", "subject", "body"],
            },
          },
          {
            name: "get_calendar_events",
            description: "Get calendar events from Outlook",
            inputSchema: {
              type: "object",
              properties: {
                start: {
                  type: "string",
                  description: "Start date (ISO format)",
                },
                end: {
                  type: "string",
                  description: "End date (ISO format)",
                },
                limit: {
                  type: "number",
                  description: "Number of events to retrieve (default: 10)",
                  default: 10,
                },
              },
            },
          },
          {
            name: "create_calendar_event",
            description: "Create a new calendar event",
            inputSchema: {
              type: "object",
              properties: {
                subject: {
                  type: "string",
                  description: "Event subject/title",
                },
                start: {
                  type: "string",
                  description: "Start date and time (ISO format)",
                },
                end: {
                  type: "string",
                  description: "End date and time (ISO format)",
                },
                attendees: {
                  type: "array",
                  items: { type: "string" },
                  description: "Email addresses of attendees",
                },
                body: {
                  type: "string",
                  description: "Event description",
                },
                location: {
                  type: "string",
                  description: "Event location",
                },
              },
              required: ["subject", "start", "end"],
            },
          },
        ],
      };
    });

    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      try {
        const { name, arguments: args } = request.params;

        if (!this.outlookClient.isAuthenticated()) {
          throw new McpError(
            ErrorCode.InvalidRequest,
            "Not authenticated with Outlook. Please authenticate first."
          );
        }

        switch (name) {
          case "get_emails":
            return await this.handleGetEmails(args);
          case "send_email":
            return await this.handleSendEmail(args);
          case "get_calendar_events":
            return await this.handleGetCalendarEvents(args);
          case "create_calendar_event":
            return await this.handleCreateCalendarEvent(args);
          default:
            throw new McpError(
              ErrorCode.MethodNotFound,
              `Unknown tool: ${name}`
            );
        }
      } catch (error) {
        if (error instanceof McpError) {
          throw error;
        }
        throw new McpError(
          ErrorCode.InternalError,
          `Tool execution failed: ${error}`
        );
      }
    });
  }

  private setupResourceHandlers() {
    this.server.setRequestHandler(ListResourcesRequestSchema, async () => {
      return {
        resources: [
          {
            uri: "outlook://profile",
            mimeType: "application/json",
            name: "Outlook Profile",
            description: "Current user's Outlook profile information",
          },
          {
            uri: "outlook://folders",
            mimeType: "application/json",
            name: "Mail Folders",
            description: "List of mail folders in Outlook",
          },
        ],
      };
    });

    this.server.setRequestHandler(ReadResourceRequestSchema, async (request) => {
      const { uri } = request.params;

      try {
        if (!this.outlookClient.isAuthenticated()) {
          throw new McpError(
            ErrorCode.InvalidRequest,
            "Not authenticated with Outlook"
          );
        }

        switch (uri) {
          case "outlook://profile":
            const profile = await this.outlookClient.getUserProfile();
            return {
              contents: [
                {
                  uri,
                  mimeType: "application/json",
                  text: JSON.stringify(profile, null, 2),
                },
              ],
            };

          case "outlook://folders":
            const folders = await this.outlookClient.getMailFolders();
            return {
              contents: [
                {
                  uri,
                  mimeType: "application/json",
                  text: JSON.stringify(folders, null, 2),
                },
              ],
            };

          default:
            throw new McpError(
              ErrorCode.InvalidRequest,
              `Unknown resource: ${uri}`
            );
        }
      } catch (error) {
        if (error instanceof McpError) {
          throw error;
        }
        throw new McpError(
          ErrorCode.InternalError,
          `Resource read failed: ${error}`
        );
      }
    });
  }

  private async handleGetEmails(args: any) {
    const { limit = 10, search, folder = "inbox" } = args;
    const emails = await this.outlookClient.getEmails(limit, search, folder);
    
    return {
      content: [
        {
          type: "text",
          text: JSON.stringify(emails, null, 2),
        },
      ],
    };
  }

  private async handleSendEmail(args: any) {
    const { to, subject, body, cc, bcc } = args;
    const result = await this.outlookClient.sendEmail(to, subject, body, cc, bcc);
    
    return {
      content: [
        {
          type: "text",
          text: `Email sent successfully. Message ID: ${result.id}`,
        },
      ],
    };
  }

  private async handleGetCalendarEvents(args: any) {
    const { start, end, limit = 10 } = args;
    const events = await this.outlookClient.getCalendarEvents(start, end, limit);
    
    return {
      content: [
        {
          type: "text",
          text: JSON.stringify(events, null, 2),
        },
      ],
    };
  }

  private async handleCreateCalendarEvent(args: any) {
    const { subject, start, end, attendees, body, location } = args;
    const event = await this.outlookClient.createCalendarEvent(
      subject,
      start,
      end,
      attendees,
      body,
      location
    );
    
    return {
      content: [
        {
          type: "text",
          text: `Calendar event created successfully. Event ID: ${event.id}`,
        },
      ],
    };
  }

  async run() {
    // Initialize Outlook client
    const config = this.loadConfig();
    this.outlookClient = new OutlookClient(config);

    // Initialize authentication only if valid credentials are provided
    if (config.clientId !== 'test-client-id' && config.clientSecret !== 'test-client-secret') {
      try {
        await this.outlookClient.initialize();
        console.error("Outlook MCP server authenticated and ready");
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : String(error);
        console.error("Warning: Authentication failed. Server will start but tools will require authentication.", errorMessage);
      }
    } else {
      console.error("Outlook MCP server running with test credentials (authentication disabled)");
    }

    const transport = new StdioServerTransport();
    await this.server.connect(transport);
    console.error("Outlook MCP server running on stdio");
  }

  private loadConfig() {
    const config = {
      clientId: process.env.OUTLOOK_CLIENT_ID || "",
      clientSecret: process.env.OUTLOOK_CLIENT_SECRET || "",
      tenantId: process.env.OUTLOOK_TENANT_ID || "common",
      redirectUri: process.env.OUTLOOK_REDIRECT_URI || "http://localhost:3000/auth/callback",
    };

    try {
      return ConfigSchema.parse(config);
    } catch (error) {
      throw new Error(
        "Invalid configuration. Please ensure OUTLOOK_CLIENT_ID and OUTLOOK_CLIENT_SECRET are set."
      );
    }
  }
}

// Start the server
const server = new OutlookMCPServer();
server.run().catch(console.error);