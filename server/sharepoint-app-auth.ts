/**
 * SharePoint Application Authentication Service
 * 
 * Uses Azure App Registration with application permissions for server-to-server access
 * No user interaction required - perfect for backend services
 */
import pino from 'pino';

// Use same logger configuration as server
const logger = pino({
  level: process.env.LOG_LEVEL || (process.env.NODE_ENV === 'production' ? 'info' : 'debug'),
  transport: process.env.NODE_ENV === 'development' ? {
    target: 'pino-pretty'
  } : undefined
});

interface AzureAppCredentials {
  tenantId: string;
  clientId: string;
  clientSecret: string;
}

export class SharePointAppAuthService {
  private accessToken: string | null = null;
  private tokenExpiry: number = 0;
  private credentials: AzureAppCredentials;

  constructor() {
    this.credentials = {
      tenantId: process.env.AZURE_TENANT_ID!,
      clientId: process.env.AZURE_CLIENT_ID!,
      clientSecret: process.env.AZURE_CLIENT_SECRET!
    };
  }

  /**
   * Get SharePoint access token using client credentials flow
   * Uses application permissions for server-to-server access
   */
  async getSharePointAccessToken(): Promise<string> {
    // Return cached token if still valid
    if (this.accessToken && Date.now() < this.tokenExpiry) {
      return this.accessToken;
    }

    const tokenUrl = `https://login.microsoftonline.com/${this.credentials.tenantId}/oauth2/v2.0/token`;
    
    const params = new URLSearchParams();
    params.append('client_id', this.credentials.clientId);
    params.append('client_secret', this.credentials.clientSecret);
    params.append('scope', 'https://graph.microsoft.com/.default');
    params.append('grant_type', 'client_credentials');

    try {
      const response = await fetch(tokenUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
        },
        body: params,
      });

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Token request failed: ${response.status} ${response.statusText} - ${errorText}`);
      }

      const tokenData = await response.json();
      
      if (!tokenData.access_token) {
        throw new Error('No access token received from Azure');
      }
      
      this.accessToken = tokenData.access_token;
      // Set expiry 5 minutes before actual expiry for safety
      this.tokenExpiry = Date.now() + (tokenData.expires_in - 300) * 1000;
      
      logger.debug('SharePoint application token acquired successfully');
      return this.accessToken!;
    } catch (error) {
      logger.error({ err: error }, 'Failed to acquire SharePoint token');
      throw error;
    }
  }

  /**
   * Make authenticated SharePoint REST API call
   */
  async makeSharePointRequest(url: string): Promise<any> {
    const token = await this.getSharePointAccessToken();
    
    const response = await fetch(url, {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json'
      }
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`SharePoint REST API error: ${response.status} ${response.statusText} - ${errorText}`);
    }

    return response.json();
  }

  isConfigured(): boolean {
    return !!(this.credentials.tenantId && this.credentials.clientId && this.credentials.clientSecret);
  }
}