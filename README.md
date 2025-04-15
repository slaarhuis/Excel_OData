# SharePoint OData Service for Templafy Integration

This service creates an OData V4 endpoint from Excel files stored in SharePoint, specifically designed to be compatible with Templafy's data connector requirements.

## Features

- Exposes Excel table data as OData V4 endpoints
- Implements Bearer token authentication required by Templafy
- Connects to SharePoint (on-premises or online) to access Excel files
- Supports Microsoft Graph API for Office 365 environments
- Provides comprehensive logging and error handling

## Prerequisites

- Node.js 16+ and npm
- SharePoint environment (on-premises or Office 365)
- Excel files with tables in SharePoint
- Templafy account with Data Connectors module enabled

## Installation

1. Clone the repository
2. Install dependencies:
   ```
   npm install
   ```
3. Create a `.env` file based on `.env.example`:
   ```
   cp .env.example .env
   ```
4. Update the `.env` file with your SharePoint and Templafy configuration
5. Build the application:
   ```
   npm run build
   ```
6. Start the server:
   ```
   npm start
   ```

## Configuration

The following environment variables need to be configured:

### Server Configuration
- `PORT`: Port number for the server (default: 3000)
- `NODE_ENV`: Environment (development, production)
- `LOG_LEVEL`: Logging level (info, warn, error, debug)

### SharePoint Authentication
- `SHAREPOINT_TENANT_ID`: Azure AD tenant ID
- `SHAREPOINT_CLIENT_ID`: Azure AD application ID
- `SHAREPOINT_CLIENT_SECRET`: Azure AD application secret
- `SHAREPOINT_SITE_URL`: SharePoint site URL

### API Authentication for Templafy
- `API_BEARER_TOKEN`: Static Bearer token for Templafy authentication
- `TOKEN_EXPIRES_IN`: Token expiration time (not used for static tokens)

### Excel File Settings
- `EXCEL_FILE_PATH`: Path to Excel file in SharePoint
- `EXCEL_TABLE_NAME`: Name of the table in Excel file

## Usage

### OData Endpoint

The OData endpoint is available at:
```
http://localhost:3000/odata
```

### Authentication

All requests to the OData endpoint must include the Bearer token in the Authorization header:
```
Authorization: Bearer your-secure-bearer-token
```

### Example Requests

1. Get metadata:
   ```
   GET http://localhost:3000/odata/$metadata
   ```

2. Get all rows:
   ```
   GET http://localhost:3000/odata/ExcelRow
   ```

3. Get a specific row:
   ```
   GET http://localhost:3000/odata/ExcelRow('1')
   ```

4. Filter rows:
   ```
   GET http://localhost:3000/odata/ExcelRow?$filter=ColumnName eq 'Value'
   ```

## Templafy Integration

To integrate with Templafy:

1. Ensure your OData service is publicly accessible via HTTPS
2. Configure Templafy to use your OData endpoint URL
3. Use the same Bearer token in both your service and Templafy configuration
4. Test the connection from Templafy

## Development

- Run in development mode:
  ```
  npm run dev
  ```

- Build the application:
  ```
  npm run build
  ```

## License

This project is licensed under the MIT License - see the LICENSE file for details.
