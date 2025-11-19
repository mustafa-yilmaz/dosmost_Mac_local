# Office 365 Word to PDF - Quick Start Guide

Get started with converting Word documents to PDFs in just 5 minutes!

## Prerequisites

Before you begin, ensure you have:

1. ✅ An Azure account (sign up at https://azure.microsoft.com/free/)
2. ✅ Admin access to create Azure AD applications
3. ✅ Word documents stored in OneDrive

## Step 1: Create Azure AD Application (5 minutes)

### 1.1 Register Application

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** > **App registrations**
3. Click **New registration**
4. Fill in the form:
   - **Name**: `Docmost Office 365 Integration`
   - **Supported account types**: Select "Accounts in any organizational directory (Any Azure AD directory - Multitenant)"
   - **Redirect URI**: Leave blank
5. Click **Register**

### 1.2 Copy Application Details

After registration, you'll see the Overview page. Copy these values:

- **Application (client) ID**: `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`
- **Directory (tenant) ID**: `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`

### 1.3 Create Client Secret

1. In the left menu, click **Certificates & secrets**
2. Click **New client secret**
3. Add a description: `Docmost Integration Secret`
4. Set expiration: `24 months` (recommended)
5. Click **Add**
6. **IMPORTANT**: Copy the secret value immediately (you won't see it again!)

### 1.4 Configure API Permissions

1. In the left menu, click **API permissions**
2. Click **Add a permission**
3. Select **Microsoft Graph**
4. Select **Application permissions**
5. Search and add these permissions:
   - ✅ `Files.ReadWrite.All`
   - ✅ `Sites.ReadWrite.All`
6. Click **Add permissions**
7. **IMPORTANT**: Click **Grant admin consent for [Your Organization]** (requires admin)
8. Confirm by clicking **Yes**

You should see green checkmarks next to both permissions.

## Step 2: Configure Application (1 minute)

### 2.1 Update Environment Variables

Add these to your `.env` file:

```bash
# Office 365 Configuration
OFFICE365_CLIENT_ID=paste-your-client-id-here
OFFICE365_CLIENT_SECRET=paste-your-client-secret-here
OFFICE365_TENANT_ID=common
```

### 2.2 Start the Application

```bash
# Start the server
pnpm run start:dev

# The Office 365 module will be automatically loaded
```

## Step 3: Test the Integration (2 minutes)

### 3.1 Health Check

```bash
curl http://localhost:3000/api/office365/health
```

Expected response:
```json
{
  "status": "ok",
  "message": "Office 365 integration is available"
}
```

### 3.2 Initialize Service

```bash
curl -X POST http://localhost:3000/api/office365/init \
  -H "Content-Type: application/json" \
  -d '{
    "clientId": "your-client-id",
    "clientSecret": "your-client-secret",
    "tenantId": "common"
  }'
```

Expected response:
```json
{
  "success": true,
  "message": "Office 365 integration initialized successfully"
}
```

### 3.3 Convert Your First Document

First, get a file ID from OneDrive:

**Option A: From OneDrive Web UI**
1. Go to https://onedrive.live.com
2. Right-click on a Word document
3. Click "Share"
4. Click "Copy link"
5. The URL will contain the file ID (after `/s!/`)

**Option B: Use the Microsoft Graph Explorer**
1. Go to https://developer.microsoft.com/graph/graph-explorer
2. Sign in with your Microsoft account
3. Run: `GET /me/drive/root/children`
4. Find your Word document and copy its `id`

Now convert it:

```bash
curl -X POST http://localhost:3000/api/office365/convert/word-to-pdf \
  -H "Content-Type: application/json" \
  -d '{
    "fileId": "YOUR_FILE_ID_HERE",
    "createShareableLink": true
  }'
```

Expected response:
```json
{
  "success": true,
  "pdfFileId": "...",
  "downloadUrl": "https://...",
  "shareableUrl": "https://1drv.ms/b/s!...",
  "metadata": {
    "filename": "document.pdf",
    "size": 245678,
    "createdDateTime": "2024-11-19T10:00:00Z",
    "lastModifiedDateTime": "2024-11-19T10:00:00Z"
  },
  "message": "Word document successfully converted to PDF"
}
```

🎉 **Congratulations!** Your Word document has been converted to PDF!

### 3.4 Access Your PDF

Copy the `shareableUrl` from the response and paste it in your browser. You'll be able to:

- ✅ View the PDF directly
- ✅ Download it
- ✅ Share it with anyone (no sign-in required!)

## Common Use Cases

### Use Case 1: Convert Multiple Documents

```bash
curl -X POST http://localhost:3000/api/office365/convert/batch \
  -H "Content-Type: application/json" \
  -d '{
    "fileIds": [
      "file-id-1",
      "file-id-2",
      "file-id-3"
    ],
    "createShareableLink": true
  }'
```

### Use Case 2: Custom Filename

```bash
curl -X POST http://localhost:3000/api/office365/convert/word-to-pdf \
  -H "Content-Type: application/json" \
  -d '{
    "fileId": "YOUR_FILE_ID",
    "customFilename": "Q4-2024-Report",
    "createShareableLink": true
  }'
```

### Use Case 3: Organization-Only Sharing

```bash
curl -X POST http://localhost:3000/api/office365/convert/word-to-pdf \
  -H "Content-Type: application/json" \
  -d '{
    "fileId": "YOUR_FILE_ID",
    "createShareableLink": true,
    "shareLinkScope": "organization"
  }'
```

## Troubleshooting

### Issue: "Failed to authenticate with Office 365"

**Solution:**
- Verify your Client ID and Secret are correct
- Check that admin consent was granted for API permissions
- Ensure the secret hasn't expired

### Issue: "File is not a Word document"

**Solution:**
- Only `.doc` and `.docx` files are supported
- Check the file extension

### Issue: "Failed to create shareable link"

**Solution:**
- Verify the `Files.ReadWrite.All` permission is granted
- Check your tenant's external sharing settings
- Ensure admin consent was granted

### Issue: "Failed to convert Word document to PDF"

**Solution:**
- Ensure the file exists and is accessible
- Check that the file is not corrupted
- Verify the file is not password-protected
- Check file size (limit is typically 10MB)

## Next Steps

- 📖 Read the full [README.md](./README.md) for detailed documentation
- 🔍 Explore [example scripts](./examples/convert-word-to-pdf.example.ts)
- 🚀 Integrate into your application workflow
- 🔐 Review security best practices

## Support

Need help?

- Check the [troubleshooting section](#troubleshooting)
- Review the [full documentation](./README.md)
- Open an issue on GitHub

## Security Best Practices

1. ✅ **Never commit secrets** to version control
2. ✅ Use environment variables for sensitive data
3. ✅ Rotate client secrets regularly
4. ✅ Use `organization` scope for internal documents
5. ✅ Monitor API usage and access logs
6. ✅ Implement rate limiting for production

---

Happy converting! 🎉
