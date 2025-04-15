# Templafy OData Authentication Requirements

## Bearer Token Authentication
- Templafy requires authentication via static Bearer Token
- Currently, Templafy doesn't support alternate authentication methods
- The Bearer token must be included in the HTTP Authorization header

## Implementation Approach
When implementing an OData V4 service for Templafy, the service needs to:

1. Accept a Bearer token in the Authorization header
2. Validate the token against a predefined static token
3. Return appropriate error responses for invalid tokens (401 Unauthorized)

## Code Example for Adding Authorization Header
When implementing the OData service, the following pattern can be used to add the Authorization header:

```csharp
// For ASP.NET Core OData service
context.SendingRequest2 += (s, e) =>
{
    e.RequestMessage.SetHeader("Authorization", token);
};
```

## Security Considerations
- Store the Bearer token securely (e.g., in environment variables or secure configuration)
- Use HTTPS for all communications (Templafy requires this)
- Consider implementing IP filtering using Templafy's static IP addresses
- Implement proper logging for authentication attempts
- Consider token rotation mechanisms if needed

## Testing Authentication
- Test the OData service with valid and invalid tokens
- Verify that proper HTTP status codes are returned
- Ensure that the service rejects requests without valid authentication

## Integration with SharePoint Authentication
When building the service, we'll need to:
1. Authenticate with SharePoint using appropriate credentials
2. Validate incoming requests from Templafy using the Bearer token
3. Map between these two authentication systems securely
