import axios, { AxiosInstance } from 'axios';
import {
  ExcelFile,
  Worksheet,
  ExcelData,
  CreateRowPayload,
  UpdateRowPayload,
  DeleteRowPayload,
} from '../types/excel.types';

class Office365Service {
  private api: AxiosInstance;
  private accessToken: string | null = null;

  constructor() {
    this.api = axios.create({
      baseURL: '/api/office365',
    });

    // Add interceptor to include MS token in headers
    this.api.interceptors.request.use((config) => {
      if (this.accessToken) {
        config.headers['x-ms-token'] = this.accessToken;
      }
      return config;
    });
  }

  setAccessToken(token: string) {
    this.accessToken = token;
  }

  async getUserDrive() {
    const response = await this.api.get('/drive');
    return response.data;
  }

  async listExcelFiles(driveId?: string): Promise<ExcelFile[]> {
    const params = driveId ? { driveId } : {};
    const response = await this.api.get('/excel/files', { params });
    return response.data;
  }

  async getWorksheets(driveId: string, itemId: string): Promise<Worksheet[]> {
    const response = await this.api.get('/excel/worksheets', {
      params: { driveId, itemId },
    });
    return response.data;
  }

  async getExcelData(
    driveId: string,
    itemId: string,
    worksheetName: string,
    range?: string,
  ): Promise<ExcelData> {
    const response = await this.api.get('/excel/data', {
      params: { driveId, itemId, worksheetName, range },
    });
    return response.data;
  }

  async createRow(payload: CreateRowPayload) {
    const response = await this.api.post('/excel/row', payload);
    return response.data;
  }

  async updateRow(payload: UpdateRowPayload) {
    const response = await this.api.put('/excel/row', payload);
    return response.data;
  }

  async deleteRow(payload: DeleteRowPayload) {
    const response = await this.api.delete('/excel/row', {
      params: payload,
    });
    return response.data;
  }
}

export const office365Service = new Office365Service();
