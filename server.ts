import express from 'express';
import cors from 'cors';
import helmet from 'helmet';
import morgan from 'morgan';
import { ODataServer } from 'odata-v4-server';
import config from './config/config';
import logger from './utils/logger';
import { authenticateToken } from './middleware/auth.middleware';
import { ExcelRowController } from './controllers/excel.controller';
import authService from './services/auth.service';
import excelService from './services/excel.service';

// Create Express app
const app = express();

// Apply middleware
app.use(helmet());
app.use(cors());
app.use(express.json());
app.use(morgan('combined'));

// Health check endpoint
app.get('/health', (req, res) => {
  res.status(200).json({ status: 'ok' });
});

// API info endpoint
app.get('/', (req, res) => {
  res.status(200).json({
    name: 'SharePoint OData Service for Templafy',
    version: '1.0.0',
    description: 'OData V4 endpoint for Excel files in SharePoint',
    endpoints: {
      health: '/health',
      odata: '/odata',
      metadata: '/odata/$metadata'
    }
  });
});

// Create OData server for Excel data
const createODataServer = (filePath: string, tableName: string) => {
  // Create controller with file path and table name
  const ExcelODataServer = ODataServer.create({
    rootPath: '/odata',
    controller: new ExcelRowController(filePath, tableName)
  });
  
  // Apply authentication middleware to OData endpoints
  app.use('/odata', authenticateToken, ExcelODataServer.handle.bind(ExcelODataServer));
  
  logger.info('OData server created successfully', { filePath, tableName });
};

// Start the server
const startServer = async () => {
  try {
    logger.info('Starting SharePoint OData Service for Templafy');
    
    // Test SharePoint authentication
    try {
      await authService.getAccessToken();
      logger.info('SharePoint authentication successful');
    } catch (error) {
      logger.warn('SharePoint authentication not configured or failed', { error });
      logger.info('Server will start, but SharePoint access will not work until configured');
    }
    
    // Get Excel file path and table name from environment variables
    const filePath = process.env.EXCEL_FILE_PATH || 'Documents/data.xlsx';
    const tableName = process.env.EXCEL_TABLE_NAME || 'Table1';
    
    // Create OData server for Excel data
    createODataServer(filePath, tableName);
    
    // Start Express server
    app.listen(config.port, () => {
      logger.info(`Server running on port ${config.port} in ${config.nodeEnv} mode`);
      logger.info(`OData endpoint available at http://localhost:${config.port}/odata`);
      logger.info(`OData metadata available at http://localhost:${config.port}/odata/$metadata`);
    });
  } catch (error) {
    logger.error('Failed to start server', { error });
    process.exit(1);
  }
};

// Handle uncaught exceptions
process.on('uncaughtException', (error) => {
  logger.error('Uncaught exception', { error });
  process.exit(1);
});

// Handle unhandled promise rejections
process.on('unhandledRejection', (reason, promise) => {
  logger.error('Unhandled promise rejection', { reason });
  process.exit(1);
});

// Start the server
startServer();

export default app;
