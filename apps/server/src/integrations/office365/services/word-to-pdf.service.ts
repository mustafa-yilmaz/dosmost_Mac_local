import { Injectable, Logger, HttpException, HttpStatus } from '@nestjs/common';
import { GraphApiService, GraphApiConfig } from './graph-api.service';
import {
  ConvertWordToPdfDto,
  ConvertWordToPdfResponseDto,
  BatchConvertWordToPdfDto,
  BatchConvertWordToPdfResponseDto,
} from '../dto/word-to-pdf.dto';

@Injectable()
export class WordToPdfService {
  private readonly logger = new Logger(WordToPdfService.name);

  constructor(private readonly graphApiService: GraphApiService) {}

  /**
   * Initialize the service with authentication
   */
  async initialize(config?: GraphApiConfig): Promise<void> {
    await this.graphApiService.initialize(config);
  }

  /**
   * Convert a single Word document to PDF and optionally create a shareable link
   */
  async convertWordToPdf(dto: ConvertWordToPdfDto, userId: string = 'me'): Promise<ConvertWordToPdfResponseDto> {
    try {
      this.logger.log(`Starting Word to PDF conversion for file: ${dto.fileId}`);

      // Step 1: Get the original file metadata
      const originalFile = await this.graphApiService.getFileMetadata(dto.fileId, userId);
      this.logger.log(`Original file: ${originalFile.name}, Size: ${originalFile.size} bytes`);

      // Step 2: Convert Word to PDF using Graph API
      const pdfBuffer = await this.graphApiService.convertWordToPdf(dto.fileId, userId);
      this.logger.log(`Conversion successful, PDF size: ${pdfBuffer.length} bytes`);

      // Step 3: Generate PDF filename
      const originalFilename = originalFile.name.replace(/\.(docx?|doc)$/i, '');
      const pdfFilename = dto.customFilename
        ? `${dto.customFilename}.pdf`
        : `${originalFilename}.pdf`;

      // Step 4: Upload PDF to OneDrive (in the same folder as the original)
      const uploadedPdf = await this.graphApiService.uploadFile(
        pdfFilename,
        pdfBuffer,
        undefined, // Will upload to root by default
        userId,
      );
      this.logger.log(`PDF uploaded to OneDrive: ${uploadedPdf.id}`);

      // Step 5: Get download URL
      const downloadUrl = await this.graphApiService.getDownloadUrl(uploadedPdf.id, userId);

      // Step 6: Create shareable link if requested
      let shareableUrl: string | undefined;
      if (dto.createShareableLink) {
        const shareLink = await this.graphApiService.createShareableLink(
          uploadedPdf.id,
          dto.shareLinkType,
          dto.shareLinkScope,
          userId,
        );
        shareableUrl = shareLink.link.webUrl;
        this.logger.log(`Shareable link created: ${shareableUrl}`);
      }

      // Step 7: Return response
      return {
        success: true,
        pdfFileId: uploadedPdf.id,
        downloadUrl: downloadUrl,
        shareableUrl: shareableUrl,
        metadata: {
          filename: uploadedPdf.name,
          size: uploadedPdf.size,
          createdDateTime: uploadedPdf.createdDateTime,
          lastModifiedDateTime: uploadedPdf.lastModifiedDateTime,
        },
        message: 'Word document successfully converted to PDF',
      };
    } catch (error) {
      this.logger.error(`Conversion failed: ${error.message}`, error.stack);
      throw new HttpException(
        `Failed to convert Word document to PDF: ${error.message}`,
        error.status || HttpStatus.INTERNAL_SERVER_ERROR,
      );
    }
  }

  /**
   * Convert multiple Word documents to PDF in batch
   */
  async batchConvertWordToPdf(
    dto: BatchConvertWordToPdfDto,
    userId: string = 'me',
  ): Promise<BatchConvertWordToPdfResponseDto> {
    this.logger.log(`Starting batch conversion of ${dto.fileIds.length} files`);

    const results = [];
    let successCount = 0;
    let failureCount = 0;

    // Process each file sequentially to avoid rate limiting
    for (const fileId of dto.fileIds) {
      try {
        const convertDto: ConvertWordToPdfDto = {
          fileId,
          createShareableLink: dto.createShareableLink,
          shareLinkType: dto.shareLinkType,
          shareLinkScope: dto.shareLinkScope,
        };

        const result = await this.convertWordToPdf(convertDto, userId);

        results.push({
          fileId,
          success: true,
          pdfFileId: result.pdfFileId,
          downloadUrl: result.downloadUrl,
          shareableUrl: result.shareableUrl,
        });

        successCount++;
        this.logger.log(`Successfully converted file ${fileId} (${successCount}/${dto.fileIds.length})`);
      } catch (error) {
        results.push({
          fileId,
          success: false,
          error: error.message,
        });

        failureCount++;
        this.logger.error(`Failed to convert file ${fileId}: ${error.message}`);
      }

      // Add a small delay between conversions to avoid rate limiting
      if (dto.fileIds.length > 1) {
        await this.delay(1000); // 1 second delay
      }
    }

    this.logger.log(`Batch conversion completed: ${successCount} succeeded, ${failureCount} failed`);

    return {
      success: failureCount === 0,
      successCount,
      failureCount,
      results,
    };
  }

  /**
   * Convert Word document by sharing URL
   */
  async convertWordToPdfByUrl(
    sharingUrl: string,
    createShareableLink: boolean = true,
    userId: string = 'me',
  ): Promise<ConvertWordToPdfResponseDto> {
    try {
      // Extract file ID from sharing URL
      // OneDrive sharing URLs typically look like: https://1drv.ms/w/s!AhKExample
      // We need to resolve this to a file ID using Graph API

      this.logger.log(`Resolving sharing URL: ${sharingUrl}`);

      // Use the shares endpoint to get the drive item
      const encodedUrl = Buffer.from(sharingUrl).toString('base64').replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');
      const shareResponse = await this.graphApiService['axiosInstance'].get(`/shares/u!${encodedUrl}/driveItem`);
      const fileId = shareResponse.data.id;

      this.logger.log(`Resolved file ID: ${fileId}`);

      // Now convert using the file ID
      return await this.convertWordToPdf(
        {
          fileId,
          createShareableLink,
        },
        userId,
      );
    } catch (error) {
      this.logger.error(`Failed to convert Word document by URL: ${error.message}`, error.stack);
      throw new HttpException(
        `Failed to convert Word document by URL: ${error.message}`,
        error.status || HttpStatus.INTERNAL_SERVER_ERROR,
      );
    }
  }

  /**
   * Helper method to add delay
   */
  private delay(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  /**
   * Validate that a file is a Word document
   */
  async validateWordDocument(fileId: string, userId: string = 'me'): Promise<boolean> {
    try {
      const metadata = await this.graphApiService.getFileMetadata(fileId, userId);
      const validExtensions = ['.doc', '.docx'];
      const hasValidExtension = validExtensions.some((ext) => metadata.name.toLowerCase().endsWith(ext));

      if (!hasValidExtension) {
        throw new HttpException(
          `File ${metadata.name} is not a Word document. Only .doc and .docx files are supported.`,
          HttpStatus.BAD_REQUEST,
        );
      }

      return true;
    } catch (error) {
      if (error instanceof HttpException) {
        throw error;
      }
      throw new HttpException(
        `Failed to validate Word document: ${error.message}`,
        error.status || HttpStatus.INTERNAL_SERVER_ERROR,
      );
    }
  }

  /**
   * Get conversion status and metadata
   */
  async getConversionInfo(pdfFileId: string, userId: string = 'me'): Promise<any> {
    try {
      const metadata = await this.graphApiService.getFileMetadata(pdfFileId, userId);
      const downloadUrl = await this.graphApiService.getDownloadUrl(pdfFileId, userId);

      return {
        fileId: metadata.id,
        filename: metadata.name,
        size: metadata.size,
        createdDateTime: metadata.createdDateTime,
        lastModifiedDateTime: metadata.lastModifiedDateTime,
        webUrl: metadata.webUrl,
        downloadUrl: downloadUrl,
      };
    } catch (error) {
      throw new HttpException(
        `Failed to get conversion info: ${error.message}`,
        error.status || HttpStatus.INTERNAL_SERVER_ERROR,
      );
    }
  }
}
