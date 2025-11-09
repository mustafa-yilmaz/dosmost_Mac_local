import { Injectable, UnauthorizedException, BadRequestException } from '@nestjs/common';
import { Client } from '@microsoft/microsoft-graph-client';
import 'isomorphic-fetch';

@Injectable()
export class GraphService {
  /**
   * Creates an authenticated Microsoft Graph client
   */
  createClient(accessToken: string): Client {
    if (!accessToken) {
      throw new UnauthorizedException('Access token is required');
    }

    return Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      },
    });
  }

  /**
   * Get Excel workbook data
   */
  async getExcelData(
    accessToken: string,
    driveId: string,
    itemId: string,
    worksheetName: string,
    range?: string,
  ) {
    try {
      const client = this.createClient(accessToken);
      const endpoint = `/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetName}/usedRange`;

      const response = await client.api(endpoint).get();

      return {
        values: response.values,
        address: response.address,
        rowCount: response.rowCount,
        columnCount: response.columnCount,
      };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      throw new BadRequestException(`Failed to get Excel data: ${errorMessage}`);
    }
  }

  /**
   * Get list of worksheets in a workbook
   */
  async getWorksheets(accessToken: string, driveId: string, itemId: string) {
    try {
      const client = this.createClient(accessToken);
      const endpoint = `/drives/${driveId}/items/${itemId}/workbook/worksheets`;

      const response = await client.api(endpoint).get();

      return response.value;
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      throw new BadRequestException(`Failed to get worksheets: ${errorMessage}`);
    }
  }

  /**
   * Create a new row in Excel worksheet
   */
  async createRow(
    accessToken: string,
    driveId: string,
    itemId: string,
    worksheetName: string,
    data: Record<string, any>,
  ) {
    try {
      const client = this.createClient(accessToken);

      // First, get the headers to map data correctly
      const headersEndpoint = `/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetName}/range(address='1:1')`;
      const headersResponse = await client.api(headersEndpoint).get();
      const headers = headersResponse.values[0];

      // Create row values in the correct order
      const rowValues = headers.map(header => data[header] || '');

      // Get the next empty row
      const usedRangeEndpoint = `/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetName}/usedRange`;
      const usedRange = await client.api(usedRangeEndpoint).get();
      const nextRow = usedRange.rowCount + 1;

      // Add the new row
      const addRowEndpoint = `/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetName}/range(address='A${nextRow}:${String.fromCharCode(64 + headers.length)}${nextRow}')`;

      const response = await client.api(addRowEndpoint).patch({
        values: [rowValues],
      });

      return {
        success: true,
        rowIndex: nextRow,
        values: rowValues,
      };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      throw new BadRequestException(`Failed to create row: ${errorMessage}`);
    }
  }

  /**
   * Update an existing row in Excel worksheet
   */
  async updateRow(
    accessToken: string,
    driveId: string,
    itemId: string,
    worksheetName: string,
    rowIndex: number,
    data: Record<string, any>,
  ) {
    try {
      const client = this.createClient(accessToken);

      // Get headers to map data correctly
      const headersEndpoint = `/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetName}/range(address='1:1')`;
      const headersResponse = await client.api(headersEndpoint).get();
      const headers = headersResponse.values[0];

      // Create row values in the correct order
      const rowValues = headers.map(header =>
        data.hasOwnProperty(header) ? data[header] : ''
      );

      // Update the row
      const updateEndpoint = `/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetName}/range(address='A${rowIndex}:${String.fromCharCode(64 + headers.length)}${rowIndex}')`;

      const response = await client.api(updateEndpoint).patch({
        values: [rowValues],
      });

      return {
        success: true,
        rowIndex,
        values: rowValues,
      };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      throw new BadRequestException(`Failed to update row: ${errorMessage}`);
    }
  }

  /**
   * Delete a row from Excel worksheet
   */
  async deleteRow(
    accessToken: string,
    driveId: string,
    itemId: string,
    worksheetName: string,
    rowIndex: number,
  ) {
    try {
      const client = this.createClient(accessToken);

      // Get the range for the row to delete
      const deleteEndpoint = `/drives/${driveId}/items/${itemId}/workbook/worksheets/${worksheetName}/range(address='${rowIndex}:${rowIndex}')`;

      await client.api(`${deleteEndpoint}/delete`).post({
        shift: 'Up',
      });

      return {
        success: true,
        deletedRowIndex: rowIndex,
      };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      throw new BadRequestException(`Failed to delete row: ${errorMessage}`);
    }
  }

  /**
   * List Excel files from OneDrive
   */
  async listExcelFiles(accessToken: string, driveId?: string) {
    try {
      const client = this.createClient(accessToken);

      let endpoint = '/me/drive/root/search(q=\'.xlsx\')';
      if (driveId) {
        endpoint = `/drives/${driveId}/root/search(q='.xlsx')`;
      }

      const response = await client.api(endpoint).get();

      return response.value;
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      throw new BadRequestException(`Failed to list Excel files: ${errorMessage}`);
    }
  }

  /**
   * Get user's OneDrive information
   */
  async getUserDrive(accessToken: string) {
    try {
      const client = this.createClient(accessToken);
      const response = await client.api('/me/drive').get();

      return response;
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      throw new BadRequestException(`Failed to get user drive: ${errorMessage}`);
    }
  }
}
