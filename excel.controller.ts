import { ODataController, Edm, odata } from 'odata-v4-server';
import { Request } from 'express';
import logger from '../utils/logger';
import excelService from '../services/excel.service';

/**
 * Entity class for Excel table rows
 * This will be dynamically extended with properties based on Excel columns
 */
@Edm.EntityType()
export class ExcelRow {
  @Edm.Key
  @Edm.Computed
  @Edm.String
  id!: string;

  // Dynamic properties will be added based on Excel columns
  [key: string]: any;
}

/**
 * OData controller for Excel data
 * Implements OData V4 protocol for Templafy compatibility
 */
@odata.type(ExcelRow)
export class ExcelRowController extends ODataController {
  private filePath: string;
  private tableName: string;
  private columns: any[] = [];

  constructor(filePath: string, tableName: string) {
    super();
    this.filePath = filePath;
    this.tableName = tableName;
    
    // Initialize columns (will be loaded on first request)
    this.initializeColumns();
  }

  /**
   * Initialize columns from Excel table
   */
  private async initializeColumns(): Promise<void> {
    try {
      // Get table columns from Excel service
      this.columns = await excelService.getTableColumns(this.filePath, this.tableName);
      logger.info('Initialized columns for OData controller', { 
        filePath: this.filePath, 
        tableName: this.tableName,
        columnCount: this.columns.length
      });
    } catch (error) {
      logger.error('Failed to initialize columns', { error });
      // We'll retry on the first request
    }
  }

  /**
   * Get all rows from Excel table
   * Implements OData GET collection operation
   */
  @odata.GET
  async getRows(@odata.context context: any, @odata.result result: any, @odata.query query: any, req: Request): Promise<ExcelRow[]> {
    try {
      logger.info('OData request: Get all rows', { 
        filePath: this.filePath, 
        tableName: this.tableName,
        query: query
      });
      
      // Ensure columns are loaded
      if (this.columns.length === 0) {
        await this.initializeColumns();
      }
      
      // Get table rows from Excel service
      const rows = await excelService.getTableRows(this.filePath, this.tableName);
      
      // Convert to OData format
      const odataRows = excelService.convertToODataFormat(rows, this.columns);
      
      // Apply OData query options (filtering, sorting, etc.)
      // This is handled automatically by the odata-v4-server library
      
      return odataRows;
    } catch (error) {
      logger.error('Error processing OData request: Get all rows', { 
        error, 
        filePath: this.filePath, 
        tableName: this.tableName 
      });
      throw error;
    }
  }

  /**
   * Get a single row by ID
   * Implements OData GET entity operation
   */
  @odata.GET
  async getRow(@odata.key key: string, @odata.context context: any, @odata.result result: any, @odata.query query: any, req: Request): Promise<ExcelRow> {
    try {
      logger.info('OData request: Get row by ID', { 
        id: key, 
        filePath: this.filePath, 
        tableName: this.tableName,
        query: query
      });
      
      // Ensure columns are loaded
      if (this.columns.length === 0) {
        await this.initializeColumns();
      }
      
      // Convert key to index (ID is the row index as string)
      const index = parseInt(key, 10);
      
      if (isNaN(index)) {
        throw new Error(`Invalid row ID: ${key}`);
      }
      
      // Get row by index
      const row = await excelService.getTableRow(this.filePath, this.tableName, index);
      
      // Convert to OData format
      const odataRows = excelService.convertToODataFormat([row], this.columns);
      
      if (odataRows.length === 0) {
        throw new Error(`Row with ID ${key} not found`);
      }
      
      return odataRows[0];
    } catch (error) {
      logger.error('Error processing OData request: Get row by ID', { 
        error, 
        id: key, 
        filePath: this.filePath, 
        tableName: this.tableName 
      });
      throw error;
    }
  }

  /**
   * Get OData metadata for Excel table
   * This is automatically handled by the odata-v4-server library
   * based on the entity type definition
   */
}
