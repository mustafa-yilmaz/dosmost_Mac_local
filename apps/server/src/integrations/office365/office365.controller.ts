import {
  Controller,
  Post,
  Get,
  Body,
  Param,
  Query,
  HttpCode,
  HttpStatus,
  UseGuards,
  Logger,
} from '@nestjs/common';
import { ApiTags, ApiOperation, ApiResponse, ApiBearerAuth, ApiParam, ApiQuery } from '@nestjs/swagger';
import { WordToPdfService } from './services/word-to-pdf.service';
import {
  ConvertWordToPdfDto,
  ConvertWordToPdfResponseDto,
  BatchConvertWordToPdfDto,
  BatchConvertWordToPdfResponseDto,
  Office365AuthDto,
} from './dto/word-to-pdf.dto';
import { JwtAuthGuard } from '../../core/auth/jwt-auth.guard';

@ApiTags('Office 365')
@Controller('office365')
export class Office365Controller {
  private readonly logger = new Logger(Office365Controller.name);

  constructor(private readonly wordToPdfService: WordToPdfService) {}

  @Post('init')
  @HttpCode(HttpStatus.OK)
  @ApiOperation({
    summary: 'Initialize Office 365 integration with credentials',
    description: 'Initialize the Office 365 service with authentication credentials (Client ID, Secret, Tenant ID)',
  })
  @ApiResponse({
    status: HttpStatus.OK,
    description: 'Successfully initialized Office 365 integration',
  })
  @ApiResponse({
    status: HttpStatus.UNAUTHORIZED,
    description: 'Authentication failed',
  })
  async initialize(@Body() authDto: Office365AuthDto): Promise<{ success: boolean; message: string }> {
    this.logger.log('Initializing Office 365 integration...');

    await this.wordToPdfService.initialize({
      clientId: authDto.clientId,
      clientSecret: authDto.clientSecret,
      tenantId: authDto.tenantId,
      accessToken: authDto.accessToken,
    });

    return {
      success: true,
      message: 'Office 365 integration initialized successfully',
    };
  }

  @Post('convert/word-to-pdf')
  @HttpCode(HttpStatus.OK)
  @ApiOperation({
    summary: 'Convert Word document to PDF',
    description:
      'Convert an online Word document from OneDrive to PDF and optionally create a public shareable link. ' +
      'The PDF will be uploaded to OneDrive and a shareable URL will be generated if requested.',
  })
  @ApiResponse({
    status: HttpStatus.OK,
    description: 'Word document successfully converted to PDF',
    type: ConvertWordToPdfResponseDto,
  })
  @ApiResponse({
    status: HttpStatus.BAD_REQUEST,
    description: 'Invalid file ID or file is not a Word document',
  })
  @ApiResponse({
    status: HttpStatus.INTERNAL_SERVER_ERROR,
    description: 'Conversion failed',
  })
  async convertWordToPdf(@Body() dto: ConvertWordToPdfDto): Promise<ConvertWordToPdfResponseDto> {
    this.logger.log(`Converting Word document to PDF: ${dto.fileId}`);

    // Validate that the file is a Word document
    await this.wordToPdfService.validateWordDocument(dto.fileId);

    // Perform conversion
    return await this.wordToPdfService.convertWordToPdf(dto);
  }

  @Post('convert/batch')
  @HttpCode(HttpStatus.OK)
  @ApiOperation({
    summary: 'Batch convert multiple Word documents to PDF',
    description:
      'Convert multiple Word documents from OneDrive to PDF in batch. ' +
      'Each file will be converted sequentially and shareable links can be created for all PDFs.',
  })
  @ApiResponse({
    status: HttpStatus.OK,
    description: 'Batch conversion completed',
    type: BatchConvertWordToPdfResponseDto,
  })
  @ApiResponse({
    status: HttpStatus.BAD_REQUEST,
    description: 'Invalid request',
  })
  async batchConvertWordToPdf(@Body() dto: BatchConvertWordToPdfDto): Promise<BatchConvertWordToPdfResponseDto> {
    this.logger.log(`Batch converting ${dto.fileIds.length} Word documents to PDF`);
    return await this.wordToPdfService.batchConvertWordToPdf(dto);
  }

  @Post('convert/by-url')
  @HttpCode(HttpStatus.OK)
  @ApiOperation({
    summary: 'Convert Word document to PDF using sharing URL',
    description:
      'Convert a Word document to PDF using a OneDrive sharing URL. ' +
      'This is useful when you have a sharing link instead of a file ID.',
  })
  @ApiResponse({
    status: HttpStatus.OK,
    description: 'Word document successfully converted to PDF',
    type: ConvertWordToPdfResponseDto,
  })
  @ApiResponse({
    status: HttpStatus.BAD_REQUEST,
    description: 'Invalid sharing URL',
  })
  async convertByUrl(
    @Body() body: { sharingUrl: string; createShareableLink?: boolean },
  ): Promise<ConvertWordToPdfResponseDto> {
    this.logger.log(`Converting Word document to PDF by URL: ${body.sharingUrl}`);
    return await this.wordToPdfService.convertWordToPdfByUrl(body.sharingUrl, body.createShareableLink);
  }

  @Get('pdf/:fileId')
  @ApiOperation({
    summary: 'Get PDF file information',
    description: 'Retrieve metadata and download information for a converted PDF file',
  })
  @ApiParam({
    name: 'fileId',
    description: 'OneDrive file ID of the PDF',
    example: '01BYE5RZ6QN3ZWBTUFOFD3GSPGOHDJD36L',
  })
  @ApiResponse({
    status: HttpStatus.OK,
    description: 'PDF file information retrieved successfully',
  })
  @ApiResponse({
    status: HttpStatus.NOT_FOUND,
    description: 'PDF file not found',
  })
  async getPdfInfo(@Param('fileId') fileId: string): Promise<any> {
    this.logger.log(`Getting PDF info: ${fileId}`);
    return await this.wordToPdfService.getConversionInfo(fileId);
  }

  @Get('health')
  @HttpCode(HttpStatus.OK)
  @ApiOperation({
    summary: 'Health check for Office 365 integration',
    description: 'Check if the Office 365 integration is properly configured and accessible',
  })
  @ApiResponse({
    status: HttpStatus.OK,
    description: 'Office 365 integration is healthy',
  })
  async healthCheck(): Promise<{ status: string; message: string }> {
    // Basic health check - just return OK
    // In a production environment, you might want to test the Graph API connection
    return {
      status: 'ok',
      message: 'Office 365 integration is available',
    };
  }
}
