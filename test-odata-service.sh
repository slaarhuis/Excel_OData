#!/bin/bash

# Script to test the OData service with Templafy requirements

# Set variables
PORT=3000
API_URL="http://localhost:$PORT"
BEARER_TOKEN="test-token"

# Create a .env file for testing
echo "Creating .env file for testing..."
cat > .env << EOL
PORT=$PORT
NODE_ENV=development
LOG_LEVEL=debug

# SharePoint Authentication (using mock values for testing)
SHAREPOINT_TENANT_ID=test-tenant-id
SHAREPOINT_CLIENT_ID=test-client-id
SHAREPOINT_CLIENT_SECRET=test-client-secret
SHAREPOINT_SITE_URL=https://test-tenant.sharepoint.com/sites/test-site

# API Authentication for Templafy
API_BEARER_TOKEN=$BEARER_TOKEN
TOKEN_EXPIRES_IN=1d

# Excel File Settings (using mock values for testing)
EXCEL_FILE_PATH=Documents/test.xlsx
EXCEL_TABLE_NAME=Table1
EOL

echo "Building the application..."
npm run build

echo "Starting the server in test mode..."
NODE_ENV=test node dist/server.js &
SERVER_PID=$!

# Wait for server to start
echo "Waiting for server to start..."
sleep 5

# Test health endpoint
echo "Testing health endpoint..."
HEALTH_RESPONSE=$(curl -s $API_URL/health)
if [[ $HEALTH_RESPONSE == *"ok"* ]]; then
  echo "✅ Health endpoint test passed"
else
  echo "❌ Health endpoint test failed"
  echo "Response: $HEALTH_RESPONSE"
fi

# Test API info endpoint
echo "Testing API info endpoint..."
INFO_RESPONSE=$(curl -s $API_URL/)
if [[ $INFO_RESPONSE == *"SharePoint OData Service"* ]]; then
  echo "✅ API info endpoint test passed"
else
  echo "❌ API info endpoint test failed"
  echo "Response: $INFO_RESPONSE"
fi

# Test OData endpoint without authentication
echo "Testing OData endpoint without authentication..."
UNAUTH_RESPONSE=$(curl -s -w "%{http_code}" $API_URL/odata/ExcelRow)
if [[ $UNAUTH_RESPONSE == *"401"* ]]; then
  echo "✅ Authentication required test passed"
else
  echo "❌ Authentication required test failed"
  echo "Response: $UNAUTH_RESPONSE"
fi

# Test OData endpoint with invalid authentication
echo "Testing OData endpoint with invalid authentication..."
INVALID_AUTH_RESPONSE=$(curl -s -w "%{http_code}" -H "Authorization: Bearer invalid-token" $API_URL/odata/ExcelRow)
if [[ $INVALID_AUTH_RESPONSE == *"403"* ]]; then
  echo "✅ Invalid authentication test passed"
else
  echo "❌ Invalid authentication test failed"
  echo "Response: $INVALID_AUTH_RESPONSE"
fi

# Test OData metadata endpoint
echo "Testing OData metadata endpoint..."
METADATA_RESPONSE=$(curl -s -H "Authorization: Bearer $BEARER_TOKEN" $API_URL/odata/\$metadata)
if [[ $METADATA_RESPONSE == *"<edmx:Edmx"* ]]; then
  echo "✅ OData metadata test passed"
else
  echo "❌ OData metadata test failed"
  echo "Response: $METADATA_RESPONSE"
fi

# Test OData entity set endpoint
echo "Testing OData entity set endpoint..."
ENTITY_SET_RESPONSE=$(curl -s -H "Authorization: Bearer $BEARER_TOKEN" $API_URL/odata/ExcelRow)
if [[ $ENTITY_SET_RESPONSE == *"\"value\":"* ]]; then
  echo "✅ OData entity set test passed"
else
  echo "❌ OData entity set test failed"
  echo "Response: $ENTITY_SET_RESPONSE"
fi

# Kill the server
echo "Stopping the server..."
kill $SERVER_PID

echo "Test completed!"
