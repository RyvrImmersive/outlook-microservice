# Outlook Attachment Microservice

A FastAPI microservice for fetching Outlook emails and downloading attachments using Microsoft Graph API with MSAL authentication.

## Features

- **Search emails** with flexible filtering (sender, subject, date range, attachments)
- **Download attachments** as base64-encoded content
- **App-only authentication** using MSAL (no user interaction required)
- **Production-ready** with proper error handling and API key authentication
- **Easy deployment** with Render.com configuration

## Setup

### 1. Azure App Registration

1. Go to [Azure Portal](https://portal.azure.com) → App registrations
2. Create new registration with these settings:
   - **Name**: Outlook Microservice
   - **Account types**: Single tenant
   - **Redirect URI**: Not needed for app-only flow
3. Note the **Application (client) ID** and **Directory (tenant) ID**
4. Go to **Certificates & secrets** → Create new client secret
5. Go to **API permissions** → Add permission:
   - **Microsoft Graph** → **Application permissions** → **Mail.Read**
   - **Grant admin consent** for the permission

### 2. Environment Variables

```bash
TENANT_ID=your-tenant-id
CLIENT_ID=your-client-id  
CLIENT_SECRET=your-client-secret
DEFAULT_USER_EMAIL=user@yourdomain.com
API_KEY=your-optional-api-key
GRAPH_BASE=https://graph.microsoft.com/v1.0
```

### 3. Local Development

```bash
# Install dependencies
pip install -r requirements.txt

# Set environment variables
export TENANT_ID="..."
export CLIENT_ID="..."
export CLIENT_SECRET="..."
export DEFAULT_USER_EMAIL="user@yourdomain.com"
export API_KEY="optional-api-key"

# Run the service
uvicorn main:app --reload --port 8000
```

## API Endpoints

### Health Check
```
GET /healthz
```

### Search Emails
```
POST /search
Content-Type: application/json
X-API-Key: your-api-key (if API_KEY is set)

{
  "user_email": "user@domain.com",
  "sender_email": "sender@domain.com",
  "subject_contains": "invoice",
  "days_back": 7,
  "top": 25,
  "folder": "inbox",
  "has_attachments": true
}
```

**Response:**
```json
{
  "items": [
    {
      "messageId": "AAMkAD...",
      "subject": "Invoice #12345",
      "from": "sender@domain.com",
      "fromName": "Sender Name",
      "receivedAt": "2025-01-10T10:30:00Z",
      "webLink": "https://outlook.office365.com/...",
      "hasAttachments": true,
      "attachments": [
        {
          "attachmentId": "AAMkAD...",
          "name": "invoice.pdf",
          "size": 245760,
          "contentType": "application/pdf"
        }
      ]
    }
  ],
  "summary": {
    "totalMessages": 1,
    "totalAttachments": 1
  },
  "debug": {
    "user": "user@domain.com",
    "folder": "inbox",
    "count": 1
  }
}
```

### Download Attachment
```
POST /download
Content-Type: application/json
X-API-Key: your-api-key (if API_KEY is set)

{
  "user_email": "user@domain.com",
  "message_id": "AAMkAD...",
  "attachment_id": "AAMkAD..."
}
```

**Response:**
```json
{
  "filename": "invoice.pdf",
  "content_type": "application/pdf",
  "size": 245760,
  "content_base64": "JVBERi0xLjQKJcfs..."
}
```

## Deployment

### Render.com (Recommended)

1. Push code to GitHub repository
2. Connect repository to Render.com
3. Use the included `render.yaml` for automatic configuration
4. Set environment variables in Render dashboard
5. Deploy!

### Manual Deployment

```bash
# Production server
uvicorn main:app --host 0.0.0.0 --port 8000 --workers 4
```

## Security Notes

- Uses **app-only authentication** (no user credentials stored)
- Requires **admin consent** for Mail.Read permission
- Optional **API key authentication** for additional security
- All Graph API calls use secure HTTPS
- Tokens are automatically refreshed by MSAL

## Integration with Langflow

This microservice can be easily integrated with Langflow workflows:

1. Use **HTTP Request** component to call `/search` endpoint
2. Parse JSON response to extract message and attachment metadata
3. Use **HTTP Request** component to call `/download` endpoint for specific attachments
4. Decode base64 content for further processing

## Troubleshooting

- **Token errors**: Verify Azure app registration and permissions
- **403 Forbidden**: Ensure admin consent is granted for Mail.Read
- **404 User not found**: Check DEFAULT_USER_EMAIL or user_email parameter
- **502 Graph errors**: Check Microsoft Graph API status and network connectivity
