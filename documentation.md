# SharePoint OData Service for Templafy Integration - Documentation

## Overview

This service creates an OData V4 endpoint from Excel files stored in SharePoint, specifically designed to be compatible with Templafy's data connector requirements. It allows Templafy to access data from Excel files in SharePoint and use it in document templates.

## Architecture

The service follows a modular architecture with clear separation of concerns:

1. **Authentication Layer**
   - SharePoint authentication using Microsoft Identity Platform
   - Templafy authentication using static Bearer token

2. **Data Access Layer**
   - Excel file access via Microsoft Graph API
   - Table, worksheet, and row operations

3. **OData Protocol Layer**
   - OData V4 implementation compatible with Templafy
   - Entity type definitions and metadata

4. **API Layer**
   - RESTful endpoints for health checks and API info
   - OData endpoints for data access

## Components

### Configuration (src/config)
- `config.ts`: Configuration module for environment variables

### Controllers (src/controllers)
- `excel.controller.ts`: OData controller for Excel data

### Middleware (src/middleware)
- `auth.middleware.ts`: Authentication middleware for Templafy

### Models (src/models)
- Entity definitions for OData protocol

### Services (src/services)
- `auth.service.ts`: Authentication service for SharePoint
- `excel.service.ts`: Excel file access service
- `sharepoint.service.ts`: SharePoint service for file operations

### Utils (src/utils)
- `logger.ts`: Logging utility

### Server (src/server.ts)
- Main application entry point

## Authentication

### SharePoint Authentication
The service uses client credentials flow to authenticate with SharePoint:
1. Requests an access token from Microsoft Identity Platform
2. Uses the token to access Excel files via Microsoft Graph API
3. Handles token caching and renewal

### Templafy Authentication
The service implements Bearer token authentication as required by Templafy:
1. Validates the Bearer token in the Authorization header
2. Returns appropriate HTTP status codes for authentication failures
3. Uses a static token as specified in Templafy's requirements

## Excel File Access

The service accesses Excel files in SharePoint using Microsoft Graph API:
1. Gets workbook metadata
2. Gets worksheets and tables
3. Gets table columns and rows
4. Converts Excel data to OData format

## OData Implementation

The service implements OData V4 protocol as required by Templafy:
1. Provides metadata endpoint (`$metadata`)
2. Supports entity set operations (GET collection)
3. Supports entity operations (GET entity by key)
4. Supports query options (filter, select, etc.)
5. Returns data in the format expected by Templafy

## Deployment

### Prerequisites
- Node.js 16+ and npm
- SharePoint environment (on-premises or Office 365)
- Excel files with tables in SharePoint
- Templafy account with Data Connectors module enabled

### Installation
1. Clone the repository
2. Install dependencies: `npm install`
3. Create a `.env` file based on `.env.example`
4. Update the `.env` file with your SharePoint and Templafy configuration
5. Build the application: `npm run build`
6. Start the server: `npm start`

### Configuration
The following environment variables need to be configured:

#### Server Configuration
- `PORT`: Port number for the server (default: 3000)
- `NODE_ENV`: Environment (development, production)
- `LOG_LEVEL`: Logging level (info, warn, error, debug)

#### SharePoint Authentication
- `SHAREPOINT_TENANT_ID`: Azure AD tenant ID
- `SHAREPOINT_CLIENT_ID`: Azure AD application ID
- `SHAREPOINT_CLIENT_SECRET`: Azure AD application secret
- `SHAREPOINT_SITE_URL`: SharePoint site URL

#### API Authentication for Templafy
- `API_BEARER_TOKEN`: Static Bearer token for Templafy authentication
- `TOKEN_EXPIRES_IN`: Token expiration time (not used for static tokens)

#### Excel File Settings
- `EXCEL_FILE_PATH`: Path to Excel file in SharePoint
- `EXCEL_TABLE_NAME`: Name of the table in Excel file

## Templafy Integration

To integrate with Templafy:

1. Ensure your OData service is publicly accessible via HTTPS
2. In Templafy Admin Center, go to Data Connectors
3. Add a new Custom Data Connector
4. Enter your OData endpoint URL (e.g., `https://your-service.com/odata`)
5. Configure the Bearer token (must match the `API_BEARER_TOKEN` in your service)
6. Test the connection
7. Map the Excel columns to Templafy fields

## Testing

The service includes comprehensive tests:
- Unit tests for individual components
- Integration tests for the OData endpoint
- Shell script for testing the service with Templafy requirements

To run the tests:
```
npm test
```

To run the manual test script:
```
chmod +x tests/test-odata-service.sh
./tests/test-odata-service.sh
```

## Troubleshooting

### Common Issues

1. **Authentication Failures**
   - Check SharePoint credentials in `.env` file
   - Verify Azure AD application permissions
   - Ensure Bearer token matches between service and Templafy

2. **Excel File Access Issues**
   - Verify file path and table name
   - Check SharePoint permissions
   - Ensure Excel file contains tables

3. **OData Protocol Issues**
   - Check OData endpoint URL
   - Verify metadata endpoint is accessible
   - Test with OData query options

### Logging

The service uses Winston for logging:
- Logs are output to console by default
- Log level can be configured via `LOG_LEVEL` environment variable
- Logs include timestamps and context information

## Security Considerations

1. **Authentication**
   - Store secrets securely (environment variables, Azure Key Vault, etc.)
   - Use HTTPS for all communications
   - Implement IP filtering if needed

2. **Data Access**
   - Limit SharePoint permissions to read-only
   - Only expose necessary data
   - Consider data sensitivity

3. **Deployment**
   - Use secure hosting environment
   - Implement proper network security
   - Keep dependencies updated

## Maintenance

1. **Updates**
   - Keep Node.js and npm packages updated
   - Monitor for security vulnerabilities
   - Update SharePoint and Microsoft Graph API versions

2. **Monitoring**
   - Implement health checks
   - Monitor API usage and performance
   - Set up alerts for failures

3. **Backup**
   - Backup configuration
   - Document deployment process
   - Maintain version control

## Conclusion

This service provides a robust solution for making Excel files hosted in SharePoint available as an OData endpoint for Templafy. It follows best practices for authentication, data access, and OData protocol implementation, ensuring compatibility with Templafy's requirements.
