# Research Notes: SharePoint OData and Templafy Integration

## SharePoint OData Integration Options

### SharePoint REST API with Excel Services
- SharePoint Server introduced REST API for Excel Workbooks stored in SharePoint document libraries
- Excel Services REST API applies to SharePoint and SharePoint 2016 on-premises
- For Office 365 Education, Business, and Enterprise accounts, use Excel REST APIs from Microsoft Graph

### OData Request Structure for SharePoint
- Service root URI: `http://<serverName>/_vti_bin/ExcelRest.aspx`
- Resource path: `/Documents/<workbookName>.xlsx/OData`
- Query options: `/Table1?$top=20` (example to get first 20 rows from Table1)

### Supported OData Query Options in SharePoint
- `<tableName>`: Returns all rows for the specified table (max 500 rows per page)
- `$metadata`: Returns all available tables and type information
- `$orderby`: Returns rows sorted by specified value
- `$top`: Returns N rows from the table
- `$skip`: Skips N rows and returns remaining rows
- `$skiptoken`: Seeks to Nth row and returns remaining rows
- `$filter`: Returns subset of rows that satisfy conditions
- `$format`: Atom XML format is the only supported value
- `$select`: Returns the specified entity
- `$inlinecount`: Returns the number of rows in the table

## Microsoft Graph API for Excel
- For Office 365 accounts, Excel REST APIs are part of Microsoft Graph
- Access workbooks through Drive API by identifying file location in URL:
  - `https://graph.microsoft.com/v1.0/me/drive/items/{id}/workbook/`
  - `https://graph.microsoft.com/v1.0/me/drive/root:/{item-path}:/workbook/`
- Supports CRUD operations on Excel objects (Table, Range, Chart)
- Only supports Office Open XML file format (.xlsx), not .xls
- Requires authentication with Microsoft identity platform
- Permission scopes needed: Files.Read (read) or Files.ReadWrite (read/write)

### Session Modes in Microsoft Graph
1. Persistent session - Changes are saved to the source location
2. Non-persistent session - Changes not saved, temporary copy used
3. Sessionless - Each operation requires locating the workbook

## Templafy OData Requirements

### General Requirements
- OData V4 connection from Templafy to client system
- HTTPS endpoint available publicly (no self-signed certificates)
- Authentication via static Bearer Token (no alternate authentication methods supported)
- OData V4 response must support Templafy's data type requirements

### Security Requirements
- Bearer token authentication
- TLS 1.2 minimum required
- Network route filtering on Templafy's static IP addresses (optional)
- No support for VPNs/VPCs, PrivateLink, or dynamic authentication tokens

### Supported Data Types
- Edm.Boolean
- Edm.Byte
- Edm.Date
- Edm.DateTimeOffset
- Edm.Decimal
- Edm.Double
- Edm.Duration
- Edm.Guid
- Edm.Int16
- Edm.Int32
- Edm.Int64
- Edm.SByte
- Edm.Single
- Edm.String

### Templafy Integration Process
1. Templafy queries OData endpoint for $metadata
2. OData server responds with entity descriptions, properties, data types, and operations
3. Templafy's OData client adapts to the metadata specification
4. Data can then be inserted into documents upon creation

## Integration Approaches

### Option 1: Direct SharePoint OData Connection (On-Premises)
- Use SharePoint's built-in Excel Services REST API with OData
- Requires SharePoint Server on-premises
- Limited to OData functionality provided by SharePoint

### Option 2: Microsoft Graph API (Office 365)
- Use Microsoft Graph API for Excel workbooks in SharePoint Online
- Requires building a middleware service to:
  - Authenticate with Microsoft Graph
  - Transform Microsoft Graph responses to OData V4 format
  - Implement Bearer Token authentication for Templafy

### Option 3: Custom OData V4 Service
- Build a custom OData V4 service that:
  - Connects to SharePoint (on-premises or online)
  - Reads Excel files from SharePoint
  - Exposes data as OData V4 endpoint
  - Implements Bearer Token authentication
  - Supports all required Templafy data types

## Next Steps
- Determine which SharePoint environment is being used (on-premises or online)
- Select the appropriate integration approach based on environment
- Set up development environment with necessary tools and libraries
- Implement authentication for both SharePoint and Templafy
- Create service that meets all Templafy OData V4 requirements
