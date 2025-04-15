import { Client } from '@microsoft/microsoft-graph-client';
import 'isomorphic-fetch';
import logger from '../utils/logger';
import authService from './auth.service';

/**
 * Service for accessing Excel files in SharePoint
 */
class ExcelService {
  private client: Client | null = null;

  /**
   * Initialize Microsoft Graph client
   */
  private async getGraphClient(): Promise<Client> {
    if (!this.client) {
      try {
        // Get access token from auth service
        const accessToken = await authService.getAccessToken();
        
        // Initialize Microsoft Graph client
        this.client = Client.init({
          authProvider: (done) => {
            done(null, accessToken);
          }
        });
        
        logger.info('Microsoft Graph client initialized successfully');
      } catch (error) {
        logger.error('Failed to initialize Microsoft Graph client', { error });
        throw new Error('Failed to initialize Microsoft Graph client');
      }
    }
    
    return this.client;
  }

  /**
   * Get Excel workbook metadata
   */
  async getWorkbookMetadata(filePath: string): Promise<any> {
    try {
      const client = await this.getGraphClient();
      
      logger.info('Fetching workbook metadata', { filePath });
      
      // Get workbook metadata
      const workbook = await client
        .api(`/sites/root/drive/root:/${filePath}:/workbook`)
        .get();
      
      return workbook;
    } catch (error) {
      logger.error('Failed to get workbook metadata', { error, filePath });
      throw new Error(`Failed to access Excel file: ${filePath}`);
    }
  }

  /**
   * Get all worksheets in workbook
   */
  async getWorksheets(filePath: string): Promise<any[]> {
    try {
      const client = await this.getGraphClient();
      
      logger.info('Fetching worksheets', { filePath });
      
      // Get all worksheets
      const response = await client
        .api(`/sites/root/drive/root:/${filePath}:/workbook/worksheets`)
        .get();
      
      return response.value;
    } catch (error) {
      logger.error('Failed to get worksheets', { error, filePath });
      throw new Error(`Failed to access worksheets in: ${filePath}`);
    }
  }

  /**
   * Get all tables in workbook
   */
  async getTables(filePath: string): Promise<any[]> {
    try {
      const client = await this.getGraphClient();
      
      logger.info('Fetching tables', { filePath });
      
      // Get all tables
      const response = await client
        .api(`/sites/root/drive/root:/${filePath}:/workbook/tables`)
        .get();
      
      return response.value;
    } catch (error) {
      logger.error('Failed to get tables', { error, filePath });
      throw new Error(`Failed to access tables in: ${filePath}`);
    }
  }

  /**
   * Get table metadata
   */
  async getTableMetadata(filePath: string, tableName: string): Promise<any> {
    try {
      const client = await this.getGraphClient();
      
      logger.info('Fetching table metadata', { filePath, tableName });
      
      // Get table metadata
      const table = await client
        .api(`/sites/root/drive/root:/${filePath}:/workbook/tables/${tableName}`)
        .get();
      
      return table;
    } catch (error) {
      logger.error('Failed to get table metadata', { error, filePath, tableName });
      throw new Error(`Failed to access table: ${tableName}`);
    }
  }

  /**
   * Get table columns
   */
  async getTableColumns(filePath: string, tableName: string): Promise<any[]> {
    try {
      const client = await this.getGraphClient();
      
      logger.info('Fetching table columns', { filePath, tableName });
      
      // Get table columns
      const response = await client
        .api(`/sites/root/drive/root:/${filePath}:/workbook/tables/${tableName}/columns`)
        .get();
      
      return response.value;
    } catch (error) {
      logger.error('Failed to get table columns', { error, filePath, tableName });
      throw new Error(`Failed to access table columns: ${tableName}`);
    }
  }

  /**
   * Get table rows
   */
  async getTableRows(filePath: string, tableName: string): Promise<any[]> {
    try {
      const client = await this.getGraphClient();
      
      logger.info('Fetching table rows', { filePath, tableName });
      
      // Get table rows
      const response = await client
        .api(`/sites/root/drive/root:/${filePath}:/workbook/tables/${tableName}/rows`)
        .get();
      
      return response.value;
    } catch (error) {
      logger.error('Failed to get table rows', { error, filePath, tableName });
      throw new Error(`Failed to access table rows: ${tableName}`);
    }
  }

  /**
   * Get table row by index
   */
  async getTableRow(filePath: string, tableName: string, index: number): Promise<any> {
    try {
      const client = await this.getGraphClient();
      
      logger.info('Fetching table row', { filePath, tableName, index });
      
      // Get table row by index
      const row = await client
        .api(`/sites/root/drive/root:/${filePath}:/workbook/tables/${tableName}/rows/itemAt(index=${index})`)
        .get();
      
      return row;
    } catch (error) {
      logger.error('Failed to get table row', { error, filePath, tableName, index });
      throw new Error(`Failed to access table row at index ${index}`);
    }
  }

  /**
   * Convert Excel data to OData compatible format
   */
  convertToODataFormat(rows: any[], columns: any[]): any[] {
    try {
      // Extract column names
      const columnNames = columns.map(col => col.name);
      
      // Convert rows to OData format
      return rows.map((row, index) => {
        const odataRow: any = {
          id: index.toString(),
        };
        
        // Add values for each column
        if (row.values && Array.isArray(row.values[0])) {
          row.values[0].forEach((value: any, colIndex: number) => {
            if (colIndex < columnNames.length) {
              const columnName = columnNames[colIndex];
              odataRow[columnName] = value;
            }
          });
        }
        
        return odataRow;
      });
    } catch (error) {
      logger.error('Failed to convert Excel data to OData format', { error });
      throw new Error('Failed to convert Excel data to OData format');
    }
  }
}

export default new ExcelService();
