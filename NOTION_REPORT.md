# Outlook MCP Server - Technical Reference Documentation

*Last Updated: December 2024*

## 📊 Project Overview

**Outlook MCP Server** is a Model Context Protocol (MCP) implementation that bridges AI assistants with Microsoft Outlook services through the Microsoft Graph API. It provides a standardized interface for comprehensive email, calendar, and contact management operations.

### Key Value Propositions
- **Standardized Integration**: MCP-compliant interface for consistent AI tool interactions
- **Comprehensive Coverage**: Full CRUD operations across Outlook services
- **Secure Authentication**: OAuth2 with automatic token refresh and secure storage
- **Production Ready**: Robust error handling, input validation, and timezone management

## 🏗️ Technical Architecture

### Core Components

#### 1. MCP Server Layer
- **Framework**: `@modelcontextprotocol/sdk` v1.0.0
- **Transport**: StdioServerTransport for AI assistant communication
- **Tool Registration**: Dynamic tool discovery and validation
- **Request Handling**: Standardized request/response formatting

#### 2. Microsoft Graph Integration
- **Client Library**: `@microsoft/microsoft-graph-client` v3.0.7
- **API Coverage**: Full Graph API v1.0 endpoint support
- **Data Transformation**: JSON serialization with proper type mapping
- **Error Handling**: Comprehensive Graph API error interpretation

#### 3. Authentication System
- **Library**: `@azure/msal-node` v2.6.4 (Microsoft Authentication Library)
- **Flow**: OAuth2 Authorization Code with PKCE
- **Token Management**: Automatic refresh, secure local caching
- **Scopes**: `https://graph.microsoft.com/.default`, `offline_access`

#### 4. Data Validation
- **Library**: Zod v3.25.51
- **Validation**: Runtime input validation and type safety
- **Error Messages**: User-friendly validation error reporting

## 🛠️ Implementation Details

### File Structure
```
├── server.js          # Main server implementation (1096 lines)
├── package.json       # Dependencies and scripts
├── .gitignore        # Git exclusions
├── .env.example      # Environment template
└── README.md         # Project documentation
```

### Environment Configuration
```env
OUTLOOK_CLIENT_ID=<Azure_App_Client_ID>
OUTLOOK_CLIENT_SECRET=<Azure_App_Client_Secret>  
OUTLOOK_TENANT_ID=<Azure_Tenant_ID>
```

### Token Caching Strategy
- **Location**: `.mcp-outlook-token-cache.json` (project root)
- **Content**: Access token, refresh token, expiration timestamp
- **Security**: File-based storage with appropriate permissions
- **Refresh Logic**: Automatic token renewal before expiration

## 🔧 Available Tools & Capabilities

### Authentication Tools
| Tool | Parameters | Description |
|------|------------|-------------|
| `outlook_auth` | `authCode?: string` | OAuth2 authentication flow handler |

### Email Management Tools
| Tool | Key Parameters | Capabilities |
|------|----------------|--------------|
| `outlook_list_emails` | `folder, limit, search` | List emails with filtering and search |
| `outlook_read_email` | `emailId` | Read full email content and metadata |
| `outlook_send_email` | `to[], subject, body, cc[], bcc[]` | Send emails with rich formatting |
| `outlook_reply_email` | `emailId, body, replyAll` | Reply to emails (single/all) |
| `outlook_forward_email` | `emailId, to[], body` | Forward emails with comments |
| `outlook_delete_email` | `emailId` | Delete emails permanently |
| `outlook_mark_email_read` | `emailId, isRead` | Mark read/unread status |
| `outlook_move_email` | `emailId, destinationFolder` | Move emails between folders |
| `outlook_list_folders` | None | List all mail folders with counts |

### Calendar Management Tools
| Tool | Key Parameters | Capabilities |
|------|----------------|--------------|
| `outlook_list_calendar_events` | `startDateTime, endDateTime, limit` | View calendar events with date filtering |
| `outlook_create_event` | `subject, start, end, location, body` | Create calendar events with full details |
| `outlook_update_event` | `eventId, updateFields` | Modify existing calendar events |
| `outlook_delete_event` | `eventId` | Remove calendar events |
| `outlook_get_event_details` | `eventId` | Get detailed event information |

### Contact Management Tools
| Tool | Key Parameters | Capabilities |
|------|----------------|--------------|
| `outlook_list_contacts` | `limit, search` | List contacts with search filtering |
| `outlook_create_contact` | `displayName, emailAddresses[]` | Create new contacts |
| `outlook_update_contact` | `contactId, updateFields` | Update contact information |
| `outlook_delete_contact` | `contactId` | Delete contacts |
| `outlook_get_contact_details` | `contactId` | Get detailed contact information |

## 🔐 Security Implementation

### Authentication Security
- **OAuth2 Best Practices**: Authorization code flow with state parameter
- **Token Security**: Encrypted token storage, automatic rotation
- **Scope Management**: Minimal required permissions
- **Tenant Isolation**: Proper tenant ID validation

### Input Validation
- **Zod Schemas**: Runtime type validation for all inputs
- **Sanitization**: Email content sanitization for XSS prevention
- **Parameter Validation**: Strict parameter type and format checking

### Error Handling
- **Security**: No sensitive data in error messages
- **Logging**: Structured error logging without credential exposure
- **Graceful Degradation**: Appropriate fallbacks for API failures

## 📈 Performance Considerations

### Token Management
- **Caching Strategy**: Local file-based token caching
- **Refresh Logic**: Proactive token refresh (5 minutes before expiry)
- **Connection Pooling**: HTTP client connection reuse

### API Optimization
- **Selective Fields**: Using `$select` to minimize data transfer
- **Pagination**: Proper handling of large result sets
- **Rate Limiting**: Built-in Graph API rate limit handling

### Timezone Handling
- **System Detection**: Automatic system timezone detection
- **Format Conversion**: Proper ISO 8601 datetime formatting
- **Graph API Compatibility**: Timezone-aware datetime objects

## 🚀 Setup & Deployment

### Prerequisites
1. **Node.js**: Version 18.0.0 or higher
2. **Azure AD App Registration**: 
   - API Permissions: `Mail.ReadWrite`, `Calendars.ReadWrite`, `Contacts.ReadWrite`
   - Authentication: Web platform with redirect URI
3. **Environment Setup**: Proper `.env` configuration

### Installation Steps
```bash
# 1. Clone repository
git clone <repository-url>
cd outlook-mcp-server

# 2. Install dependencies
npm install

# 3. Configure environment
cp .env.example .env
# Edit .env with your Azure credentials

# 4. Start server
npm start

# 5. Authenticate
# Use outlook_auth tool to complete OAuth flow
```

### Deployment Considerations
- **Environment Variables**: Secure credential management
- **File Permissions**: Proper token cache file permissions
- **Network Access**: Ensure Graph API endpoint accessibility
- **Logging**: Configure appropriate log levels for production

## 🔄 Integration Patterns

### MCP Client Integration
```javascript
// Example client-side tool usage
const result = await mcpClient.callTool('outlook_send_email', {
  to: ['recipient@example.com'],
  subject: 'Test Email',
  body: '<h1>Hello World</h1>',
  importance: 'high'
});
```

### Error Handling Pattern
```javascript
try {
  const emails = await mcpClient.callTool('outlook_list_emails', {
    folder: 'inbox',
    limit: 50,
    search: 'important'
  });
} catch (error) {
  // Handle authentication, permission, or API errors
  console.error('Email listing failed:', error.message);
}
```

## 📊 Monitoring & Observability

### Key Metrics to Track
- **Authentication Success Rate**: OAuth flow completion rate
- **API Response Times**: Graph API call latency
- **Error Rates**: By tool and error type
- **Token Refresh Frequency**: Authentication health indicator

### Logging Strategy
- **Structured Logging**: JSON format with correlation IDs
- **Security**: No PII or credentials in logs
- **Performance**: Request/response timing
- **Errors**: Full error context without sensitive data

## 🔮 Future Enhancement Opportunities

### High Priority
1. **Rate Limiting**: Implement client-side rate limiting
2. **Caching Layer**: Redis-based caching for frequently accessed data
3. **Unit Testing**: Comprehensive test suite with mocks
4. **Webhook Support**: Real-time event notifications

### Medium Priority
1. **Batch Operations**: Multi-email operations support
2. **Advanced Search**: Complex query building
3. **Attachment Handling**: File upload/download capabilities
4. **Teams Integration**: Expand to Microsoft Teams API

### Low Priority
1. **Multi-tenant Support**: Support multiple Office 365 tenants
2. **Monitoring Dashboard**: Real-time usage analytics
3. **Plugin Architecture**: Extensible tool framework
4. **Configuration UI**: Web-based setup interface

## 📋 Troubleshooting Guide

### Common Issues
1. **Authentication Failures**: Check client ID, secret, and tenant ID
2. **Permission Errors**: Verify Graph API permissions in Azure AD
3. **Token Expiry**: Ensure automatic refresh is working
4. **Network Issues**: Check Graph API endpoint accessibility

### Debug Steps
1. Enable verbose logging
2. Verify environment variables
3. Test Graph API access directly
4. Check token cache file permissions
5. Validate Azure AD app configuration

---

*This documentation serves as a comprehensive technical reference for the Outlook MCP Server implementation. For quick start information, refer to the main README.md file.* 