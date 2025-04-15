import { Client } from '@microsoft/microsoft-graph-client';
import 'isomorphic-fetch';
import config from '../config/config';
import logger from '../utils/logger';

/**
 * Service for authenticating with SharePoint and accessing Excel files
 */
class SharePointService {
  private client: Client | null = null;

  /**
   * Initialize Microsoft Graph client with SharePoint credentials
   */
  async initialize(): Promise<void> {
    try {
      // Get access token for Microsoft Graph API
      const accessToken = await this.getAccessToken();
      
      // Initialize Microsoft Graph client
      this.client = Client.init({
        authProvider: (done) => {
          done(null, accessToken);
        }
      });
      
      logger.info('SharePoint service initialized successfully');
    } catch (error) {
      logger.error('Failed to initialize SharePoint service', { error });
      throw new Error('SharePoint authentication failed');
    }
  }

  /**
   * Get access token for Microsoft Graph API using client credentials flow
   */
  private async getAccessToken(): Promise<string> {
    try {
      const { tenantId, clientId, clientSecret } = config.sharepoint;
      
      // Validate required configuration
      if (!tenantId || !clientId || !clientSecret) {
        throw new Error('Missing SharePoint authentication configuration');
      }

      // For demonstration purposes - in a real implementation, this would use
      // MSAL or similar library to get a token from Azure AD
      const tokenEndpoint = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
      
      const response = await fetch(tokenEndpoint, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
        },
        body: new URLSearchParams({
          client_id: clientId,
          client_secret: clientSecret,
          scope: 'https://graph.microsoft.com/.default',
          grant_type: 'client_credentials',
        }),
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(`Token request failed: ${errorData.error_description || 'Unknown error'}`);
      }

      const data = await response.json();
      return data.access_token;
    } catch (error) {
      logger.error('Failed to get access token', { error });
      throw new Error('SharePoint authentication failed');
    }
  }

  /**
   * Get Excel workbook from SharePoint
   */
  async getWorkbook(filePath: string): Promise<any> {
    try {
      if (!this.client) {
        await this.initialize();
      }

      if (!this.client) {
        throw new Error('Microsoft Graph client not initialized');
      }

      // Get workbook from SharePoint
      const workbook = await this.client
        .api(`/sites/root/drive/root:/${filePath}:/workbook`)
        .get();

      return workbook;
    } catch (error) {
      logger.error('Failed to get workbook', { error, filePath });
      throw new Error(`Failed to access Excel file: ${filePath}`);
    }
  }

  /**
   * Get worksheet from workbook
   */
  async getWorksheet(filePath: string, worksheetName: string): Promise<any> {
    try {
      if (!this.client) {
        await this.initialize();
      }

      if (!this.client) {
        throw new Error('Microsoft Graph client not initialized');
      }

      // Get worksheet from workbook
      const worksheet = await this.client
        .api(`/sites/root/drive/root:/${filePath}:/workbook/worksheets/${worksheetName}`)
        .get();

      return worksheet;
    } catch (error) {
      logger.error('Failed to get worksheet', { error, filePath, worksheetName });
      throw new Error(`Failed to access worksheet: ${worksheetName}`);
    }
  }

  /**
   * Get table from worksheet
   */
  async getTable(filePath: string, tableName: string): Promise<any> {
    try {
      if (!this.client) {
        await this.initialize();
      }

      if (!this.client) {
        throw new Error('Microsoft Graph client not initialized');
      }

      // Get table from workbook
      const table = await this.client
        .api(`/sites/root/drive/root:/${filePath}:/workbook/tables/${tableName}`)
        .get();

      return table;
    } catch (error) {
      logger.error('Failed to get table', { error, filePath, tableName });
      throw new Error(`Failed to access table: ${tableName}`);
    }
  }

  /**
   * Get table rows
   */
  async getTableRows(filePath: string, tableName: string): Promise<any[]> {
    try {
      if (!this.client) {
        await this.initialize();
      }

      if (!this.client) {
        throw new Error('Microsoft Graph client not initialized');
      }

      // Get table rows from workbook
      const response = await this.client
        .api(`/sites/root/drive/root:/${filePath}:/workbook/tables/${tableName}/rows`)
        .get();

      return response.value;
    } catch (error) {
      logger.error('Failed to get table rows', { error, filePath, tableName });
      throw new Error(`Failed to access table rows: ${tableName}`);
    }
  }
}

export default new SharePointService();
