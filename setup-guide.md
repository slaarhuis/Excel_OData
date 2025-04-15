# SharePoint OData Service for Templafy Integration - Setup Guide

## Prerequisites

Before setting up the service, ensure you have:

1. **Node.js Environment**
   - Node.js 16+ installed
   - npm 7+ installed

2. **SharePoint Environment**
   - SharePoint Online (Office 365) or SharePoint Server on-premises
   - Excel files with tables stored in SharePoint
   - Admin access to register an application in Azure AD (for SharePoint Online)

3. **Templafy Account**
   - Templafy tenant with admin access
   - Data Connectors module enabled
   - Custom Data Connector feature enabled

## Step 1: Register an Application in Azure AD (for SharePoint Online)

1. Sign in to the [Azure Portal](https://portal.azure.com)
2. Navigate to Azure Active Directory > App registrations
3. Click "New registration"
4. Enter a name for your application (e.g., "SharePoint OData Service")
5. Select "Accounts in this organizational directory only"
6. Click "Register"
7. Note the Application (client) ID and Directory (tenant) ID
8. Navigate to "Certificates & secrets"
9. Create a new client secret and note the value
10. Navigate to "API permissions"
11. Add the following permissions:
    - Microsoft Graph > Application permissions > Files.Read.All
    - Microsoft Graph > Application permissions > Sites.Read.All
12. Click "Grant admin consent"

## Step 2: Set Up the Service

1. Clone the repository:
   ```
   git clone https://github.com/your-org/sharepoint-odata-service.git
   cd sharepoint-odata-service
   ```

2. Install dependencies:
   ```
   npm install
   ```

3. Create a `.env` file:
   ```
   cp .env.example .env
   ```

4. Update the `.env` file with your configuration:
   ```
   PORT=3000
   NODE_ENV=production
   LOG_LEVEL=info

   # SharePoint Authentication
   SHAREPOINT_TENANT_ID=your-tenant-id
   SHAREPOINT_CLIENT_ID=your-client-id
   SHAREPOINT_CLIENT_SECRET=your-client-secret
   SHAREPOINT_SITE_URL=https://your-tenant.sharepoint.com/sites/your-site

   # API Authentication for Templafy
   API_BEARER_TOKEN=your-secure-bearer-token
   TOKEN_EXPIRES_IN=1d

   # Excel File Settings
   EXCEL_FILE_PATH=Documents/your-excel-file.xlsx
   EXCEL_TABLE_NAME=Table1
   ```

5. Build the application:
   ```
   npm run build
   ```

6. Start the server:
   ```
   npm start
   ```

## Step 3: Deploy the Service

### Option 1: Deploy to a Node.js Hosting Service

1. Choose a hosting service (e.g., Azure App Service, Heroku, AWS Elastic Beanstalk)
2. Follow the hosting service's deployment instructions
3. Set environment variables in the hosting service's configuration
4. Ensure the service is accessible via HTTPS

### Option 2: Deploy Using Docker

1. Build the Docker image:
   ```
   docker build -t sharepoint-odata-service .
   ```

2. Run the Docker container:
   ```
   docker run -p 3000:3000 --env-file .env sharepoint-odata-service
   ```

3. Deploy to a container orchestration service (e.g., Kubernetes, Azure Container Instances)

## Step 4: Configure Templafy

1. Sign in to Templafy Admin Center
2. Navigate to Data Connectors
3. Click "Add Data Connector"
4. Select "Custom Data Connector"
5. Enter the OData endpoint URL (e.g., `https://your-service.com/odata`)
6. Enter the Bearer token (must match the `API_BEARER_TOKEN` in your service)
7. Test the connection
8. Map the Excel columns to Templafy fields
9. Save the configuration

## Step 5: Test the Integration

1. Create a new document template in Templafy
2. Add data fields from your Custom Data Connector
3. Test the template to ensure data is being pulled correctly
4. Publish the template for users

## Troubleshooting

If you encounter issues:

1. Check the service logs for errors
2. Verify all environment variables are set correctly
3. Ensure the Excel file and table exist in SharePoint
4. Test the OData endpoint directly using a tool like Postman
5. Verify the Bearer token matches between the service and Templafy

## Security Recommendations

1. Use a strong, randomly generated Bearer token
2. Store secrets securely (environment variables, Azure Key Vault, etc.)
3. Use HTTPS for all communications
4. Implement IP filtering to restrict access to Templafy's IP addresses
5. Regularly rotate credentials and update dependencies

## Support

For support, please contact:
- Email: support@your-company.com
- Internal ticketing system: [link]
- Documentation: [link]
