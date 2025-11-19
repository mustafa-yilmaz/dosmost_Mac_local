# Office 365 Integration - Word to PDF Converter

This module provides seamless integration with Microsoft Office 365 to convert Word documents (.doc, .docx) stored in OneDrive to PDF format with public shareable links, without requiring users to sign in to Microsoft Office 365.

## Features

- ✅ **Convert Word to PDF**: Leverage Microsoft Graph API to convert Word documents to PDF format
- ✅ **Public Shareable Links**: Generate anonymous shareable links for PDFs without sign-in requirements
- ✅ **Batch Conversion**: Convert multiple Word documents in a single request
- ✅ **URL-based Conversion**: Support conversion using OneDrive sharing URLs
- ✅ **Automatic Authentication**: Handle OAuth 2.0 client credentials flow automatically
- ✅ **File Validation**: Ensure only valid Word documents are processed
- ✅ **Error Handling**: Comprehensive error handling with detailed logging

## Prerequisites

Before using this integration, you need to set up an Azure AD application:

### 1. Create Azure AD Application

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** > **App registrations**
3. Click **New registration**
4. Enter a name (e.g., "Docmost Office 365 Integration")
5. Select **Accounts in any organizational directory (Any Azure AD directory - Multitenant)**
6. Click **Register**

### 2. Configure API Permissions

1. In your app registration, go to **API permissions**
2. Click **Add a permission** > **Microsoft Graph** > **Application permissions**
3. Add the following permissions:
   - `Files.ReadWrite.All` - Read and write files in all site collections
   - `Sites.ReadWrite.All` - Read and write items in all site collections
4. Click **Grant admin consent** (requires admin privileges)

### 3. Create Client Secret

1. Go to **Certificates & secrets**
2. Click **New client secret**
3. Add a description and set expiration
4. Click **Add**
5. **Copy the secret value immediately** (you won't be able to see it again)

### 4. Get Application Details

1. Go to **Overview** section
2. Copy the following values:
   - **Application (client) ID**
   - **Directory (tenant) ID**

## Installation

### 1. Install Dependencies

The axios package is required for HTTP requests:

```bash
cd apps/server
pnpm add axios
```

### 2. Configure Environment Variables

Add the following to your `.env` file:

```bash
# Office 365 / Microsoft Graph API Configuration
OFFICE365_CLIENT_ID=your-client-id-here
OFFICE365_CLIENT_SECRET=your-client-secret-here
OFFICE365_TENANT_ID=common
```

**Note**: Use `common` for the tenant ID to support multi-tenant applications, or specify your organization's tenant ID for single-tenant.

## API Endpoints

### 1. Initialize Integration

Initialize the Office 365 service with authentication credentials:

**POST** `/api/office365/init`

**Request Body:**
```json
{
  "clientId": "your-client-id",
  "clientSecret": "your-client-secret",
  "tenantId": "common"
}
```

**Response:**
```json
{
  "success": true,
  "message": "Office 365 integration initialized successfully"
}
```

### 2. Convert Word to PDF

Convert a single Word document from OneDrive to PDF:

**POST** `/api/office365/convert/word-to-pdf`

**Request Body:**
```json
{
  "fileId": "01BYE5RZ6QN3ZWBTUFOFD3GSPGOHDJD36K",
  "createShareableLink": true,
  "shareLinkType": "view",
  "shareLinkScope": "anonymous",
  "customFilename": "my-converted-document"
}
```

**Request Parameters:**
- `fileId` (required): OneDrive file ID of the Word document
- `createShareableLink` (optional, default: true): Create a public shareable link
- `shareLinkType` (optional, default: "view"): Type of link - "view", "edit", or "embed"
- `shareLinkScope` (optional, default: "anonymous"): Scope - "anonymous" or "organization"
- `customFilename` (optional): Custom filename for the PDF (without extension)

**Response:**
```json
{
  "success": true,
  "pdfFileId": "01BYE5RZ6QN3ZWBTUFOFD3GSPGOHDJD36L",
  "downloadUrl": "https://graph.microsoft.com/v1.0/me/drive/items/01BYE5RZ6QN3ZWBTUFOFD3GSPGOHDJD36L/content",
  "shareableUrl": "https://1drv.ms/b/s!AhKExample",
  "metadata": {
    "filename": "document.pdf",
    "size": 245678,
    "createdDateTime": "2024-11-19T10:00:00Z",
    "lastModifiedDateTime": "2024-11-19T10:00:00Z"
  },
  "message": "Word document successfully converted to PDF"
}
```

### 3. Batch Convert Multiple Documents

Convert multiple Word documents in a single request:

**POST** `/api/office365/convert/batch`

**Request Body:**
```json
{
  "fileIds": [
    "01BYE5RZ6QN3ZWBTUFOFD3GSPGOHDJD36K",
    "01BYE5RZ6QN3ZWBTUFOFD3GSPGOHDJD36M"
  ],
  "createShareableLink": true,
  "shareLinkType": "view",
  "shareLinkScope": "anonymous"
}
```

**Response:**
```json
{
  "success": true,
  "successCount": 2,
  "failureCount": 0,
  "results": [
    {
      "fileId": "01BYE5RZ6QN3ZWBTUFOFD3GSPGOHDJD36K",
      "success": true,
      "pdfFileId": "01BYE5RZ6QN3ZWBTUFOFD3GSPGOHDJD36L",
      "downloadUrl": "https://...",
      "shareableUrl": "https://1drv.ms/b/s!AhKExample1"
    },
    {
      "fileId": "01BYE5RZ6QN3ZWBTUFOFD3GSPGOHDJD36M",
      "success": true,
      "pdfFileId": "01BYE5RZ6QN3ZWBTUFOFD3GSPGOHDJD36N",
      "downloadUrl": "https://...",
      "shareableUrl": "https://1drv.ms/b/s!AhKExample2"
    }
  ]
}
```

### 4. Convert by Sharing URL

Convert a Word document using a OneDrive sharing URL:

**POST** `/api/office365/convert/by-url`

**Request Body:**
```json
{
  "sharingUrl": "https://1drv.ms/w/s!AhKExample",
  "createShareableLink": true
}
```

**Response:** Same as single conversion endpoint

### 5. Get PDF Information

Retrieve metadata for a converted PDF:

**GET** `/api/office365/pdf/:fileId`

**Response:**
```json
{
  "fileId": "01BYE5RZ6QN3ZWBTUFOFD3GSPGOHDJD36L",
  "filename": "document.pdf",
  "size": 245678,
  "createdDateTime": "2024-11-19T10:00:00Z",
  "lastModifiedDateTime": "2024-11-19T10:00:00Z",
  "webUrl": "https://onedrive.live.com/...",
  "downloadUrl": "https://graph.microsoft.com/v1.0/..."
}
```

### 6. Health Check

Check if the Office 365 integration is available:

**GET** `/api/office365/health`

**Response:**
```json
{
  "status": "ok",
  "message": "Office 365 integration is available"
}
```

## Usage Examples

### Using cURL

```bash
# Initialize the integration
curl -X POST http://localhost:3000/api/office365/init \
  -H "Content-Type: application/json" \
  -d '{
    "clientId": "your-client-id",
    "clientSecret": "your-client-secret",
    "tenantId": "common"
  }'

# Convert a Word document to PDF
curl -X POST http://localhost:3000/api/office365/convert/word-to-pdf \
  -H "Content-Type: application/json" \
  -d '{
    "fileId": "01BYE5RZ6QN3ZWBTUFOFD3GSPGOHDJD36K",
    "createShareableLink": true
  }'
```

### Using JavaScript/TypeScript

```typescript
import axios from 'axios';

// Initialize the service
await axios.post('http://localhost:3000/api/office365/init', {
  clientId: process.env.OFFICE365_CLIENT_ID,
  clientSecret: process.env.OFFICE365_CLIENT_SECRET,
  tenantId: 'common',
});

// Convert Word to PDF
const response = await axios.post('http://localhost:3000/api/office365/convert/word-to-pdf', {
  fileId: '01BYE5RZ6QN3ZWBTUFOFD3GSPGOHDJD36K',
  createShareableLink: true,
  shareLinkType: 'view',
  shareLinkScope: 'anonymous',
});

console.log('PDF URL:', response.data.shareableUrl);
console.log('Download URL:', response.data.downloadUrl);
```

### Using Python

```python
import requests

# Initialize the service
init_response = requests.post(
    'http://localhost:3000/api/office365/init',
    json={
        'clientId': 'your-client-id',
        'clientSecret': 'your-client-secret',
        'tenantId': 'common'
    }
)

# Convert Word to PDF
convert_response = requests.post(
    'http://localhost:3000/api/office365/convert/word-to-pdf',
    json={
        'fileId': '01BYE5RZ6QN3ZWBTUFOFD3GSPGOHDJD36K',
        'createShareableLink': True
    }
)

pdf_data = convert_response.json()
print(f"PDF URL: {pdf_data['shareableUrl']}")
```

## How It Works

### Authentication Flow

1. The service uses **OAuth 2.0 Client Credentials Flow** for authentication
2. Access tokens are automatically managed and refreshed when expired
3. No user interaction or sign-in is required

### Conversion Process

1. **Validate**: Verify the file is a valid Word document (.doc or .docx)
2. **Convert**: Use Microsoft Graph API's built-in conversion (`?format=pdf`)
3. **Upload**: Save the PDF to OneDrive (in the root folder by default)
4. **Share**: Create an anonymous shareable link (if requested)
5. **Return**: Provide file ID, download URL, and shareable link

### Shareable Link Types

- **view**: Read-only access (default, most secure for sharing)
- **edit**: Allows editing (not recommended for anonymous sharing)
- **embed**: Embeddable link for iframes

### Shareable Link Scopes

- **anonymous**: Anyone with the link can access (no sign-in required)
- **organization**: Only users in your organization can access

## Security Considerations

1. **Client Secret Protection**: Never expose client secrets in client-side code
2. **Scope Limitation**: Grant only necessary API permissions
3. **Link Expiration**: Consider implementing link expiration for sensitive documents
4. **Access Logging**: Monitor API usage and access patterns
5. **Rate Limiting**: Implement rate limiting to prevent abuse

## Troubleshooting

### Common Errors

**Error: "Office 365 credentials not configured"**
- Solution: Ensure `OFFICE365_CLIENT_ID` and `OFFICE365_CLIENT_SECRET` are set in `.env`

**Error: "Failed to authenticate with Office 365"**
- Solution: Verify your client ID, secret, and tenant ID are correct
- Check that admin consent has been granted for API permissions

**Error: "File is not a Word document"**
- Solution: Only `.doc` and `.docx` files are supported for conversion

**Error: "Failed to convert Word document to PDF"**
- Solution: Check that the file exists and is accessible
- Ensure the file is not corrupted or password-protected

**Error: "Failed to create shareable link"**
- Solution: Verify the `Files.ReadWrite.All` permission is granted
- Check that external sharing is enabled in your tenant settings

### Enable Detailed Logging

The service includes comprehensive logging. Check your application logs for detailed error messages:

```bash
# View logs in development
pnpm run start:dev

# The service logs will show:
# - Authentication attempts
# - Conversion progress
# - File operations
# - Error details
```

## Architecture

### Module Structure

```
apps/server/src/integrations/office365/
├── dto/
│   └── word-to-pdf.dto.ts          # Request/response DTOs
├── services/
│   ├── graph-api.service.ts        # Microsoft Graph API client
│   └── word-to-pdf.service.ts      # Conversion orchestration
├── office365.controller.ts         # API endpoints
├── office365.module.ts             # NestJS module
├── index.ts                        # Exports
└── README.md                       # Documentation
```

### Key Components

1. **GraphApiService**: Low-level Microsoft Graph API client
   - Handles authentication and token management
   - Provides methods for file operations
   - Manages API requests and error handling

2. **WordToPdfService**: High-level conversion orchestration
   - Validates Word documents
   - Orchestrates conversion workflow
   - Handles batch operations
   - Generates shareable links

3. **Office365Controller**: REST API endpoints
   - Exposes HTTP endpoints for conversions
   - Validates request DTOs
   - Returns formatted responses

## Limitations

- **File Size**: Microsoft Graph API has file size limits (typically 10MB for conversions)
- **Rate Limiting**: Microsoft Graph API has rate limits (consider implementing queuing for large batches)
- **File Location**: PDFs are uploaded to the root folder by default (consider adding folder selection)
- **User Context**: Uses application permissions (me endpoint may not work with client credentials)

## Future Enhancements

- [ ] Support for other Office formats (Excel, PowerPoint)
- [ ] Custom folder selection for PDF output
- [ ] Webhook support for asynchronous conversions
- [ ] Conversion progress tracking
- [ ] PDF optimization options (compression, quality)
- [ ] Conversion history and audit logs
- [ ] Support for user delegated permissions
- [ ] Queue integration for batch processing

## API Reference

For detailed API documentation, run the application and visit:
```
http://localhost:3000/api/docs
```

The Swagger/OpenAPI documentation provides interactive API testing and detailed schema information.

## License

This module is part of the Docmost project and follows the same license.

## Support

For issues, questions, or contributions, please visit the main repository or contact the development team.
