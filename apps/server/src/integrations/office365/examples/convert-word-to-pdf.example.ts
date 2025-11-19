/**
 * Office 365 Word to PDF Conversion - Usage Examples
 *
 * This file demonstrates how to use the Office 365 integration
 * to convert Word documents to PDF format with public shareable links.
 */

import axios from 'axios';

// Configuration
const API_BASE_URL = 'http://localhost:3000/api/office365';
const OFFICE365_CLIENT_ID = process.env.OFFICE365_CLIENT_ID || 'your-client-id';
const OFFICE365_CLIENT_SECRET = process.env.OFFICE365_CLIENT_SECRET || 'your-client-secret';
const OFFICE365_TENANT_ID = process.env.OFFICE365_TENANT_ID || 'common';

// Example 1: Initialize the Office 365 service
async function initializeService() {
  console.log('Initializing Office 365 service...');

  try {
    const response = await axios.post(`${API_BASE_URL}/init`, {
      clientId: OFFICE365_CLIENT_ID,
      clientSecret: OFFICE365_CLIENT_SECRET,
      tenantId: OFFICE365_TENANT_ID,
    });

    console.log('✅ Service initialized:', response.data);
    return true;
  } catch (error) {
    console.error('❌ Failed to initialize:', error.response?.data || error.message);
    return false;
  }
}

// Example 2: Convert a single Word document to PDF
async function convertSingleDocument(fileId: string) {
  console.log(`\nConverting Word document: ${fileId}`);

  try {
    const response = await axios.post(`${API_BASE_URL}/convert/word-to-pdf`, {
      fileId: fileId,
      createShareableLink: true,
      shareLinkType: 'view',
      shareLinkScope: 'anonymous',
      customFilename: 'my-converted-document',
    });

    const result = response.data;
    console.log('✅ Conversion successful!');
    console.log('PDF File ID:', result.pdfFileId);
    console.log('Download URL:', result.downloadUrl);
    console.log('Shareable URL:', result.shareableUrl);
    console.log('File Size:', (result.metadata.size / 1024).toFixed(2), 'KB');

    return result;
  } catch (error) {
    console.error('❌ Conversion failed:', error.response?.data || error.message);
    throw error;
  }
}

// Example 3: Convert multiple documents in batch
async function batchConvertDocuments(fileIds: string[]) {
  console.log(`\nBatch converting ${fileIds.length} documents...`);

  try {
    const response = await axios.post(`${API_BASE_URL}/convert/batch`, {
      fileIds: fileIds,
      createShareableLink: true,
      shareLinkType: 'view',
      shareLinkScope: 'anonymous',
    });

    const result = response.data;
    console.log('✅ Batch conversion completed!');
    console.log(`Success: ${result.successCount}, Failed: ${result.failureCount}`);

    result.results.forEach((item, index) => {
      if (item.success) {
        console.log(`  [${index + 1}] ✅ ${item.fileId}`);
        console.log(`      Shareable URL: ${item.shareableUrl}`);
      } else {
        console.log(`  [${index + 1}] ❌ ${item.fileId}: ${item.error}`);
      }
    });

    return result;
  } catch (error) {
    console.error('❌ Batch conversion failed:', error.response?.data || error.message);
    throw error;
  }
}

// Example 4: Convert using a OneDrive sharing URL
async function convertByShareUrl(sharingUrl: string) {
  console.log(`\nConverting document from sharing URL...`);

  try {
    const response = await axios.post(`${API_BASE_URL}/convert/by-url`, {
      sharingUrl: sharingUrl,
      createShareableLink: true,
    });

    const result = response.data;
    console.log('✅ Conversion successful!');
    console.log('PDF File ID:', result.pdfFileId);
    console.log('Shareable URL:', result.shareableUrl);

    return result;
  } catch (error) {
    console.error('❌ Conversion failed:', error.response?.data || error.message);
    throw error;
  }
}

// Example 5: Get PDF information
async function getPdfInfo(pdfFileId: string) {
  console.log(`\nGetting PDF information: ${pdfFileId}`);

  try {
    const response = await axios.get(`${API_BASE_URL}/pdf/${pdfFileId}`);

    const info = response.data;
    console.log('✅ PDF Information:');
    console.log('  Filename:', info.filename);
    console.log('  Size:', (info.size / 1024).toFixed(2), 'KB');
    console.log('  Created:', new Date(info.createdDateTime).toLocaleString());
    console.log('  Web URL:', info.webUrl);

    return info;
  } catch (error) {
    console.error('❌ Failed to get PDF info:', error.response?.data || error.message);
    throw error;
  }
}

// Example 6: Health check
async function checkHealth() {
  console.log('Checking Office 365 integration health...');

  try {
    const response = await axios.get(`${API_BASE_URL}/health`);
    console.log('✅ Health status:', response.data);
    return response.data.status === 'ok';
  } catch (error) {
    console.error('❌ Health check failed:', error.response?.data || error.message);
    return false;
  }
}

// Main execution
async function main() {
  console.log('='.repeat(60));
  console.log('Office 365 Word to PDF Conversion - Examples');
  console.log('='.repeat(60));

  // Step 1: Health check
  const isHealthy = await checkHealth();
  if (!isHealthy) {
    console.error('\n❌ Service is not healthy. Exiting...');
    return;
  }

  // Step 2: Initialize service
  const isInitialized = await initializeService();
  if (!isInitialized) {
    console.error('\n❌ Failed to initialize service. Exiting...');
    return;
  }

  // Step 3: Example conversions (uncomment to use)

  // Convert a single document
  // Replace with your actual OneDrive file ID
  // await convertSingleDocument('01BYE5RZ6QN3ZWBTUFOFD3GSPGOHDJD36K');

  // Convert multiple documents
  // await batchConvertDocuments([
  //   '01BYE5RZ6QN3ZWBTUFOFD3GSPGOHDJD36K',
  //   '01BYE5RZ6QN3ZWBTUFOFD3GSPGOHDJD36M',
  // ]);

  // Convert using a sharing URL
  // await convertByShareUrl('https://1drv.ms/w/s!AhKExample');

  // Get PDF information
  // await getPdfInfo('01BYE5RZ6QN3ZWBTUFOFD3GSPGOHDJD36L');

  console.log('\n' + '='.repeat(60));
  console.log('Examples completed!');
  console.log('='.repeat(60));
}

// Run examples if this file is executed directly
if (require.main === module) {
  main().catch((error) => {
    console.error('Fatal error:', error);
    process.exit(1);
  });
}

// Export functions for use in other modules
export {
  initializeService,
  convertSingleDocument,
  batchConvertDocuments,
  convertByShareUrl,
  getPdfInfo,
  checkHealth,
};
