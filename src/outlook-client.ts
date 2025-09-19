import { ConfidentialClientApplication } from "@azure/msal-node";
import { Client } from "@microsoft/microsoft-graph-client";
import { AuthenticationProvider } from "@microsoft/microsoft-graph-client";

interface OutlookConfig {
  clientId: string;
  clientSecret: string;
  tenantId: string;
  redirectUri?: string;
}

class CustomAuthProvider implements AuthenticationProvider {
  private accessToken: string | null = null;

  constructor(private cca: ConfidentialClientApplication) {}

  async getAccessToken(): Promise<string> {
    if (!this.accessToken) {
      const clientCredentialRequest = {
        scopes: ['https://graph.microsoft.com/.default'],
      };

      try {
        // For client credentials flow, use acquireTokenByClientCredential directly
        const response = await this.cca.acquireTokenByClientCredential(clientCredentialRequest);
        this.accessToken = response?.accessToken || null;
      } catch (error) {
        throw new Error(`Failed to acquire access token: ${error}`);
      }
    }

    if (!this.accessToken) {
      throw new Error('Failed to acquire access token');
    }

    return this.accessToken;
  }
}

export class OutlookClient {
  private cca: ConfidentialClientApplication;
  private graphClient: Client | null = null;
  private authProvider: CustomAuthProvider;

  constructor(private config: OutlookConfig) {
    this.cca = new ConfidentialClientApplication({
      auth: {
        clientId: config.clientId,
        clientSecret: config.clientSecret,
        authority: `https://login.microsoftonline.com/${config.tenantId}`,
      },
    });

    this.authProvider = new CustomAuthProvider(this.cca);
  }

  async initialize(): Promise<void> {
    try {
      // Initialize Graph client with custom auth provider
      this.graphClient = Client.initWithMiddleware({
        authProvider: this.authProvider,
      });

      // Test authentication by making a simple request
      await this.graphClient.api('/me').get();
    } catch (error) {
      console.error('Failed to initialize Outlook client:', error);
      throw new Error('Authentication failed. Please check your credentials.');
    }
  }

  isAuthenticated(): boolean {
    return this.graphClient !== null;
  }

  async getUserProfile(): Promise<any> {
    if (!this.graphClient) {
      throw new Error('Not authenticated');
    }

    try {
      const user = await this.graphClient.api('/me').get();
      return {
        id: user.id,
        displayName: user.displayName,
        mail: user.mail,
        userPrincipalName: user.userPrincipalName,
        jobTitle: user.jobTitle,
        department: user.department,
        officeLocation: user.officeLocation,
      };
    } catch (error) {
      throw new Error(`Failed to get user profile: ${error}`);
    }
  }

  async getEmails(limit = 10, search?: string, folder = 'inbox'): Promise<any[]> {
    if (!this.graphClient) {
      throw new Error('Not authenticated');
    }

    try {
      let query = this.graphClient
        .api(`/me/mailFolders/${folder}/messages`)
        .top(limit)
        .orderby('receivedDateTime desc')
        .select('id,subject,from,receivedDateTime,bodyPreview,isRead,importance');

      if (search) {
        query = query.search(search);
      }

      const response = await query.get();
      return response.value || [];
    } catch (error) {
      throw new Error(`Failed to get emails: ${error}`);
    }
  }

  async sendEmail(
    to: string[],
    subject: string,
    body: string,
    cc?: string[],
    bcc?: string[]
  ): Promise<any> {
    if (!this.graphClient) {
      throw new Error('Not authenticated');
    }

    try {
      const message = {
        subject,
        body: {
          contentType: 'HTML',
          content: body,
        },
        toRecipients: to.map(email => ({
          emailAddress: { address: email },
        })),
        ccRecipients: cc?.map(email => ({
          emailAddress: { address: email },
        })) || [],
        bccRecipients: bcc?.map(email => ({
          emailAddress: { address: email },
        })) || [],
      };

      const response = await this.graphClient
        .api('/me/sendMail')
        .post({ message });

      return { id: 'sent', success: true };
    } catch (error) {
      throw new Error(`Failed to send email: ${error}`);
    }
  }

  async getCalendarEvents(start?: string, end?: string, limit = 10): Promise<any[]> {
    if (!this.graphClient) {
      throw new Error('Not authenticated');
    }

    try {
      let query = this.graphClient
        .api('/me/events')
        .top(limit)
        .orderby('start/dateTime')
        .select('id,subject,start,end,location,attendees,body');

      if (start && end) {
        const startDate = new Date(start).toISOString();
        const endDate = new Date(end).toISOString();
        query = query.filter(`start/dateTime ge '${startDate}' and end/dateTime le '${endDate}'`);
      }

      const response = await query.get();
      return response.value || [];
    } catch (error) {
      throw new Error(`Failed to get calendar events: ${error}`);
    }
  }

  async createCalendarEvent(
    subject: string,
    start: string,
    end: string,
    attendees?: string[],
    body?: string,
    location?: string
  ): Promise<any> {
    if (!this.graphClient) {
      throw new Error('Not authenticated');
    }

    try {
      const event = {
        subject,
        start: {
          dateTime: new Date(start).toISOString(),
          timeZone: 'UTC',
        },
        end: {
          dateTime: new Date(end).toISOString(),
          timeZone: 'UTC',
        },
        body: body ? {
          contentType: 'HTML',
          content: body,
        } : undefined,
        location: location ? { displayName: location } : undefined,
        attendees: attendees?.map(email => ({
          emailAddress: { address: email, name: email },
          type: 'required',
        })) || [],
      };

      const response = await this.graphClient
        .api('/me/events')
        .post(event);

      return response;
    } catch (error) {
      throw new Error(`Failed to create calendar event: ${error}`);
    }
  }

  async getMailFolders(): Promise<any[]> {
    if (!this.graphClient) {
      throw new Error('Not authenticated');
    }

    try {
      const response = await this.graphClient
        .api('/me/mailFolders')
        .select('id,displayName,childFolderCount,unreadItemCount,totalItemCount')
        .get();

      return response.value || [];
    } catch (error) {
      throw new Error(`Failed to get mail folders: ${error}`);
    }
  }
}