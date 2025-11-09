# Office 365 Excel CRUD Application Setup Guide

This guide will help you set up the Office 365 Excel CRUD application that allows users to perform Create, Read, Update, and Delete operations on Excel Online files using Microsoft Graph API.

## Features

- **Microsoft 365 Authentication**: Secure OAuth 2.0 authentication using MSAL (Microsoft Authentication Library)
- **Excel File Management**: Browse and select Excel files from OneDrive
- **Real-time CRUD Operations**:
  - **Create**: Add new rows to Excel worksheets
  - **Read**: View data from Excel worksheets in a table format
  - **Update**: Edit existing rows
  - **Delete**: Remove rows from worksheets
- **Multi-worksheet Support**: Select from different worksheets within a workbook
- **User-friendly Interface**: Clean, responsive UI built with React

## Prerequisites

1. An Azure AD (Microsoft Entra ID) account
2. Access to Azure Portal (https://portal.azure.com)
3. A Microsoft 365 account with OneDrive access
4. Node.js and pnpm installed

## Azure AD Application Setup

### Step 1: Register a New Application

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** > **App registrations**
3. Click **New registration**
4. Fill in the application details:
   - **Name**: `Docmost Office 365 Excel Integration` (or your preferred name)
   - **Supported account types**: Choose based on your needs:
     - *Accounts in this organizational directory only* (Single tenant)
     - *Accounts in any organizational directory* (Multi-tenant)
     - *Accounts in any organizational directory and personal Microsoft accounts* (Recommended for broader access)
   - **Redirect URI**:
     - Platform: `Single-page application (SPA)`
     - URI: `http://localhost:5173/office365/callback` (for development)
     - For production, add: `https://your-domain.com/office365/callback`

5. Click **Register**

### Step 2: Configure API Permissions

1. In your app registration, go to **API permissions**
2. Click **Add a permission**
3. Select **Microsoft Graph**
4. Choose **Delegated permissions**
5. Add the following permissions:
   - `User.Read` - Sign in and read user profile
   - `Files.ReadWrite` - Have full access to user files
   - `Files.ReadWrite.All` - Have full access to all files user can access
6. Click **Add permissions**
7. Click **Grant admin consent** (if you have admin rights, otherwise ask your admin)

### Step 3: Configure Authentication

1. Go to **Authentication** in your app registration
2. Under **Single-page application**, ensure your redirect URI is listed
3. Under **Implicit grant and hybrid flows**, enable:
   - ✅ Access tokens (used for implicit flows)
   - ✅ ID tokens (used for implicit and hybrid flows)
4. Under **Allow public client flows**, select **No**
5. Click **Save**

### Step 4: Get Your Client ID

1. Go to **Overview** in your app registration
2. Copy the **Application (client) ID**
3. You'll need this for the environment configuration

## Application Configuration

### Backend Configuration

The backend is already configured and doesn't require additional environment variables. The Microsoft Graph API calls are made server-side using the access token passed from the frontend.

### Frontend Configuration

1. Create a `.env` file in the `apps/client` directory (or update existing one):

```bash
VITE_MICROSOFT_CLIENT_ID=your_application_client_id_here
```

2. Replace `your_application_client_id_here` with the Client ID from Azure Portal

## Running the Application

### Development Mode

1. Start the backend server:
```bash
pnpm run server:dev
```

2. Start the frontend development server:
```bash
pnpm run client:dev
```

3. Navigate to `http://localhost:5173/office365/excel`

### Production Mode

1. Build the application:
```bash
pnpm run build
```

2. Start the production server:
```bash
pnpm run start
```

## Usage Guide

### Authentication

1. Navigate to `/office365/excel`
2. Click **Sign in with Microsoft**
3. You'll be redirected to Microsoft login
4. Grant the requested permissions
5. You'll be redirected back to the application

### Working with Excel Files

#### Viewing Data

1. After signing in, you'll see a list of Excel files from your OneDrive
2. Click on a file to select it
3. Choose a worksheet from the selected file
4. The data will be displayed in a table format

#### Creating Rows

1. Click the **Add Row** button
2. Fill in the data for each column
3. Click **Save** to add the row to Excel
4. The table will refresh automatically

#### Updating Rows

1. Click the **Edit** button on the row you want to update
2. Modify the data in the input fields
3. Click **Save** to update the row
4. Click **Cancel** to discard changes

#### Deleting Rows

1. Click the **Delete** button on the row you want to remove
2. Confirm the deletion
3. The row will be removed from Excel and the table will refresh

## Architecture

### Backend (NestJS)

- **Module**: `apps/server/src/integrations/office365/`
- **Controller**: `office365.controller.ts` - Handles HTTP requests
- **Service**: `services/graph.service.ts` - Interacts with Microsoft Graph API
- **DTOs**: `dto/excel.dto.ts` - Data transfer objects for validation

### Frontend (React)

- **Components**: `apps/client/src/features/office365/components/`
  - `excel-crud.tsx` - Main component
  - `excel-file-picker.tsx` - File and worksheet selection
  - `excel-data-grid.tsx` - Data table with CRUD operations
- **Hooks**: `apps/client/src/features/office365/hooks/`
  - `use-microsoft-auth.ts` - Authentication logic
  - `use-excel-data.ts` - Data fetching and CRUD operations
- **Services**: `apps/client/src/features/office365/services/`
  - `office365-service.ts` - API client
- **Config**: `apps/client/src/features/office365/config/`
  - `msal-config.ts` - MSAL configuration

## API Endpoints

All endpoints require authentication via JWT and Microsoft access token in `x-ms-token` header.

- `GET /api/office365/drive` - Get user's OneDrive information
- `GET /api/office365/excel/files` - List Excel files
- `GET /api/office365/excel/worksheets` - Get worksheets in a workbook
- `GET /api/office365/excel/data` - Read Excel data
- `POST /api/office365/excel/row` - Create a new row
- `PUT /api/office365/excel/row` - Update an existing row
- `DELETE /api/office365/excel/row` - Delete a row

## Security Considerations

1. **Access Tokens**: Never expose access tokens in client-side code or logs
2. **HTTPS**: Always use HTTPS in production
3. **Token Storage**: Tokens are stored in sessionStorage for security
4. **Permissions**: Request only the minimum required permissions
5. **Validation**: All inputs are validated on the backend
6. **Authentication**: JWT authentication is required for all API endpoints

## Troubleshooting

### "Access token is required" Error

- Make sure you're signed in with Microsoft
- Check that the access token is being passed in the `x-ms-token` header

### "Failed to get Excel data" Error

- Verify the file permissions in OneDrive
- Ensure the Excel file is not corrupted
- Check that the worksheet name is correct

### Authentication Popup Blocked

- Allow popups for your domain in browser settings
- Try using a different browser

### CORS Errors

- Ensure your redirect URI is correctly configured in Azure AD
- Check that the domain matches exactly (including http/https)

## Additional Resources

- [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/)
- [MSAL.js Documentation](https://github.com/AzureAD/microsoft-authentication-library-for-js)
- [Excel API in Microsoft Graph](https://docs.microsoft.com/en-us/graph/api/resources/excel)
- [Azure AD App Registration Guide](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)

## Support

For issues or questions:
1. Check the Azure AD app configuration
2. Review browser console for error messages
3. Check server logs for backend errors
4. Verify API permissions in Azure Portal

## License

This feature is part of Docmost and follows the same license.
