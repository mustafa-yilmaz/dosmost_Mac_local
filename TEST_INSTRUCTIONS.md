# Testing Office 365 Excel CRUD Application

## Quick Start

### 1. Start the servers (in separate terminals):

**Terminal 1 - Backend:**
```bash
pnpm run server:dev
```

**Terminal 2 - Frontend:**
```bash
pnpm run client:dev
```

### 2. Access the Application

Open your browser and navigate to:
```
http://localhost:5173/office365/excel
```

## Testing Scenarios

### Scenario 1: Authentication
1. Click **"Sign in with Microsoft"**
2. You'll be redirected to Microsoft login
3. Sign in with your Microsoft 365 account
4. Grant the requested permissions when prompted
5. You should be redirected back and see the Excel file picker

**Expected Result:** You're authenticated and can see your name in the header

### Scenario 2: Browse Excel Files
1. After signing in, you should see a list of Excel files from your OneDrive
2. If you don't have Excel files, create a test file:
   - Go to [OneDrive](https://onedrive.live.com)
   - Create a new Excel file
   - Add some headers in row 1: `Name`, `Email`, `Status`
   - Add a few sample rows of data
   - Save the file
   - Refresh the application

**Expected Result:** Excel files are displayed with names and last modified dates

### Scenario 3: Select Worksheet
1. Click on an Excel file from the list
2. You should see a list of worksheets (tabs) from that file
3. Click on a worksheet name

**Expected Result:** Data from the worksheet is displayed in a table format

### Scenario 4: Create a New Row
1. After selecting a worksheet, click **"Add Row"** button
2. Fill in the data for each column
3. Click **"Save"**

**Expected Result:**
- The new row appears in the table
- The data is also added to the actual Excel file in OneDrive
- You can verify by opening the file in Excel Online

### Scenario 5: Update a Row
1. Find a row you want to edit
2. Click the **"Edit"** button
3. Modify the data in the input fields
4. Click **"Save"**

**Expected Result:**
- The row is updated in the table
- Changes are reflected in the Excel file
- Click "Cancel" to discard changes

### Scenario 6: Delete a Row
1. Find a row you want to delete
2. Click the **"Delete"** button
3. Confirm the deletion in the prompt

**Expected Result:**
- The row is removed from the table
- The row is deleted from the Excel file
- Remaining rows shift up

### Scenario 7: Switch Between Files
1. Select a different Excel file
2. Choose a worksheet
3. Perform CRUD operations

**Expected Result:** You can work with multiple files seamlessly

### Scenario 8: Sign Out
1. Click **"Sign Out"** in the header
2. You should be signed out

**Expected Result:** You're redirected to the login screen

## Troubleshooting During Testing

### Issue: "Access token is required" error
**Solution:**
- Make sure you're signed in
- Try signing out and signing back in
- Check browser console for errors

### Issue: No Excel files showing
**Solution:**
- Verify you have Excel files in your OneDrive
- Check the OneDrive permissions
- Ensure API permissions are granted in Azure AD

### Issue: Authentication popup blocked
**Solution:**
- Allow popups for localhost:5173
- Try a different browser (Chrome recommended)

### Issue: CORS errors
**Solution:**
- Verify the redirect URI in Azure AD matches exactly: `http://localhost:5173/office365/callback`
- Check that it's configured as "Single-page application" type

### Issue: "Failed to get Excel data"
**Solution:**
- Ensure the Excel file is not corrupted
- Verify the file has at least one row with headers
- Check that the worksheet name is correct

## Sample Test Data

If you need test data, here's a sample Excel structure:

**Headers (Row 1):**
| Name | Email | Role | Status |
|------|-------|------|--------|

**Sample Data:**
| John Doe | john@example.com | Developer | Active |
| Jane Smith | jane@example.com | Designer | Active |
| Bob Johnson | bob@example.com | Manager | Inactive |

## API Testing (Optional)

You can also test the backend APIs directly using curl or Postman:

### Get Excel Files
```bash
curl -X GET "http://localhost:3000/api/office365/excel/files" \
  -H "Authorization: Bearer YOUR_JWT_TOKEN" \
  -H "x-ms-token: YOUR_MS_ACCESS_TOKEN"
```

### Get Excel Data
```bash
curl -X GET "http://localhost:3000/api/office365/excel/data?driveId=DRIVE_ID&itemId=ITEM_ID&worksheetName=Sheet1" \
  -H "Authorization: Bearer YOUR_JWT_TOKEN" \
  -H "x-ms-token: YOUR_MS_ACCESS_TOKEN"
```

### Create Row
```bash
curl -X POST "http://localhost:3000/api/office365/excel/row" \
  -H "Content-Type: application/json" \
  -H "Authorization: Bearer YOUR_JWT_TOKEN" \
  -H "x-ms-token: YOUR_MS_ACCESS_TOKEN" \
  -d '{
    "driveId": "DRIVE_ID",
    "itemId": "ITEM_ID",
    "worksheetName": "Sheet1",
    "data": {
      "Name": "Test User",
      "Email": "test@example.com",
      "Role": "Tester",
      "Status": "Active"
    }
  }'
```

## Browser Console Debugging

Open browser DevTools (F12) and check:

1. **Console Tab:** Look for authentication errors or API failures
2. **Network Tab:** Inspect API calls to see request/response details
3. **Application Tab → Session Storage:** Check if MSAL tokens are stored

## Verification Checklist

- [ ] Can sign in with Microsoft 365
- [ ] Excel files list loads correctly
- [ ] Can select an Excel file
- [ ] Worksheets list appears
- [ ] Can select a worksheet
- [ ] Data displays in table format
- [ ] Can add a new row
- [ ] New row appears in Excel file (verify in Excel Online)
- [ ] Can edit an existing row
- [ ] Changes persist in Excel file
- [ ] Can delete a row
- [ ] Row is removed from Excel file
- [ ] Can sign out successfully

## Performance Testing

Test with different file sizes:
- Small file: < 100 rows
- Medium file: 100-1000 rows
- Large file: 1000+ rows (may be slow due to API limits)

**Note:** Microsoft Graph API has rate limits, so very large files may take time to load.

## Next Steps After Testing

Once testing is complete:
1. Update the redirect URI in Azure AD for production
2. Set production environment variables
3. Build and deploy the application
4. Test in production environment

## Need Help?

- Check `OFFICE365_SETUP.md` for detailed setup instructions
- Review browser console for error messages
- Verify Azure AD configuration
- Check server logs for backend errors
