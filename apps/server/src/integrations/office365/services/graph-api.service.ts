import { Injectable, Logger, HttpException, HttpStatus } from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import axios, { AxiosInstance, AxiosRequestConfig } from 'axios';

export interface GraphApiConfig {
  clientId: string;
  clientSecret: string;
  tenantId: string;
  accessToken?: string;
}

export interface DriveItem {
  id: string;
  name: string;
  size: number;
  createdDateTime: string;
  lastModifiedDateTime: string;
  webUrl?: string;
  '@microsoft.graph.downloadUrl'?: string;
}

export interface ShareLink {
  id: string;
  link: {
    type: string;
    scope: string;
    webUrl: string;
  };
}

@Injectable()
export class GraphApiService {
  private readonly logger = new Logger(GraphApiService.name);
  private axiosInstance: AxiosInstance;
  private accessToken: string | null = null;
  private tokenExpiration: Date | null = null;

  constructor(private configService: ConfigService) {
    this.axiosInstance = axios.create({
      baseURL: 'https://graph.microsoft.com/v1.0',
      timeout: 60000, // 60 seconds
      headers: {
        'Content-Type': 'application/json',
      },
    });

    // Add request interceptor for automatic token injection
    this.axiosInstance.interceptors.request.use(
      async (config) => {
        await this.ensureValidToken();
        if (this.accessToken) {
          config.headers.Authorization = `Bearer ${this.accessToken}`;
        }
        return config;
      },
      (error) => {
        return Promise.reject(error);
      },
    );

    // Add response interceptor for error handling
    this.axiosInstance.interceptors.response.use(
      (response) => response,
      (error) => {
        this.logger.error(
          `Graph API Error: ${error.response?.status} - ${error.response?.data?.error?.message || error.message}`,
        );
        return Promise.reject(error);
      },
    );
  }

  /**
   * Initialize the service with authentication credentials
   */
  async initialize(config?: GraphApiConfig): Promise<void> {
    if (config?.accessToken) {
      this.accessToken = config.accessToken;
      this.logger.log('Initialized with provided access token');
      return;
    }

    const clientId = config?.clientId || this.configService.get<string>('OFFICE365_CLIENT_ID');
    const clientSecret = config?.clientSecret || this.configService.get<string>('OFFICE365_CLIENT_SECRET');
    const tenantId = config?.tenantId || this.configService.get<string>('OFFICE365_TENANT_ID') || 'common';

    if (!clientId || !clientSecret) {
      throw new HttpException(
        'Office 365 credentials not configured. Please set OFFICE365_CLIENT_ID and OFFICE365_CLIENT_SECRET.',
        HttpStatus.INTERNAL_SERVER_ERROR,
      );
    }

    await this.authenticate(clientId, clientSecret, tenantId);
  }

  /**
   * Authenticate using client credentials flow
   */
  private async authenticate(clientId: string, clientSecret: string, tenantId: string): Promise<void> {
    try {
      this.logger.log('Authenticating with Microsoft Graph API...');

      const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
      const params = new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        scope: 'https://graph.microsoft.com/.default',
        grant_type: 'client_credentials',
      });

      const response = await axios.post(tokenUrl, params, {
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
        },
      });

      this.accessToken = response.data.access_token;
      const expiresIn = response.data.expires_in; // seconds
      this.tokenExpiration = new Date(Date.now() + expiresIn * 1000);

      this.logger.log('Successfully authenticated with Microsoft Graph API');
    } catch (error) {
      this.logger.error('Authentication failed:', error.response?.data || error.message);
      throw new HttpException(
        `Failed to authenticate with Office 365: ${error.response?.data?.error_description || error.message}`,
        HttpStatus.UNAUTHORIZED,
      );
    }
  }

  /**
   * Ensure the access token is still valid, refresh if needed
   */
  private async ensureValidToken(): Promise<void> {
    if (!this.accessToken || (this.tokenExpiration && new Date() >= this.tokenExpiration)) {
      this.logger.log('Token expired or not available, re-authenticating...');
      await this.initialize();
    }
  }

  /**
   * Get file metadata from OneDrive
   */
  async getFileMetadata(fileId: string, userId: string = 'me'): Promise<DriveItem> {
    try {
      const response = await this.axiosInstance.get(`/users/${userId}/drive/items/${fileId}`);
      return response.data;
    } catch (error) {
      throw new HttpException(
        `Failed to get file metadata: ${error.response?.data?.error?.message || error.message}`,
        error.response?.status || HttpStatus.INTERNAL_SERVER_ERROR,
      );
    }
  }

  /**
   * Download file content from OneDrive
   */
  async downloadFile(fileId: string, userId: string = 'me'): Promise<Buffer> {
    try {
      const response = await this.axiosInstance.get(`/users/${userId}/drive/items/${fileId}/content`, {
        responseType: 'arraybuffer',
      });
      return Buffer.from(response.data);
    } catch (error) {
      throw new HttpException(
        `Failed to download file: ${error.response?.data?.error?.message || error.message}`,
        error.response?.status || HttpStatus.INTERNAL_SERVER_ERROR,
      );
    }
  }

  /**
   * Upload file to OneDrive
   */
  async uploadFile(
    fileName: string,
    fileContent: Buffer,
    parentFolderId?: string,
    userId: string = 'me',
  ): Promise<DriveItem> {
    try {
      const uploadUrl = parentFolderId
        ? `/users/${userId}/drive/items/${parentFolderId}:/${fileName}:/content`
        : `/users/${userId}/drive/root:/${fileName}:/content`;

      const response = await this.axiosInstance.put(uploadUrl, fileContent, {
        headers: {
          'Content-Type': 'application/pdf',
        },
      });

      return response.data;
    } catch (error) {
      throw new HttpException(
        `Failed to upload file: ${error.response?.data?.error?.message || error.message}`,
        error.response?.status || HttpStatus.INTERNAL_SERVER_ERROR,
      );
    }
  }

  /**
   * Convert Word document to PDF using Microsoft Graph API
   */
  async convertWordToPdf(fileId: string, userId: string = 'me'): Promise<Buffer> {
    try {
      this.logger.log(`Converting Word document ${fileId} to PDF...`);

      // Microsoft Graph API endpoint for format conversion
      const response = await this.axiosInstance.get(`/users/${userId}/drive/items/${fileId}/content?format=pdf`, {
        responseType: 'arraybuffer',
        timeout: 120000, // 2 minutes for conversion
      });

      this.logger.log(`Successfully converted Word document to PDF`);
      return Buffer.from(response.data);
    } catch (error) {
      this.logger.error(`Failed to convert Word to PDF: ${error.response?.data?.error?.message || error.message}`);
      throw new HttpException(
        `Failed to convert Word document to PDF: ${error.response?.data?.error?.message || error.message}`,
        error.response?.status || HttpStatus.INTERNAL_SERVER_ERROR,
      );
    }
  }

  /**
   * Create a shareable link for a file
   */
  async createShareableLink(
    fileId: string,
    type: 'view' | 'edit' | 'embed' = 'view',
    scope: 'anonymous' | 'organization' = 'anonymous',
    userId: string = 'me',
  ): Promise<ShareLink> {
    try {
      this.logger.log(`Creating shareable link for file ${fileId}...`);

      const response = await this.axiosInstance.post(`/users/${userId}/drive/items/${fileId}/createLink`, {
        type: type,
        scope: scope,
      });

      this.logger.log(`Successfully created shareable link`);
      return response.data;
    } catch (error) {
      this.logger.error(`Failed to create shareable link: ${error.response?.data?.error?.message || error.message}`);
      throw new HttpException(
        `Failed to create shareable link: ${error.response?.data?.error?.message || error.message}`,
        error.response?.status || HttpStatus.INTERNAL_SERVER_ERROR,
      );
    }
  }

  /**
   * Delete a file from OneDrive
   */
  async deleteFile(fileId: string, userId: string = 'me'): Promise<void> {
    try {
      await this.axiosInstance.delete(`/users/${userId}/drive/items/${fileId}`);
      this.logger.log(`Successfully deleted file ${fileId}`);
    } catch (error) {
      this.logger.error(`Failed to delete file: ${error.response?.data?.error?.message || error.message}`);
      throw new HttpException(
        `Failed to delete file: ${error.response?.data?.error?.message || error.message}`,
        error.response?.status || HttpStatus.INTERNAL_SERVER_ERROR,
      );
    }
  }

  /**
   * Search for files in OneDrive
   */
  async searchFiles(query: string, userId: string = 'me'): Promise<DriveItem[]> {
    try {
      const response = await this.axiosInstance.get(`/users/${userId}/drive/root/search(q='${query}')`);
      return response.data.value || [];
    } catch (error) {
      throw new HttpException(
        `Failed to search files: ${error.response?.data?.error?.message || error.message}`,
        error.response?.status || HttpStatus.INTERNAL_SERVER_ERROR,
      );
    }
  }

  /**
   * Get direct download URL for a file
   */
  async getDownloadUrl(fileId: string, userId: string = 'me'): Promise<string> {
    try {
      const metadata = await this.getFileMetadata(fileId, userId);
      return metadata['@microsoft.graph.downloadUrl'] || metadata.webUrl || '';
    } catch (error) {
      throw new HttpException(
        `Failed to get download URL: ${error.response?.data?.error?.message || error.message}`,
        error.response?.status || HttpStatus.INTERNAL_SERVER_ERROR,
      );
    }
  }
}
