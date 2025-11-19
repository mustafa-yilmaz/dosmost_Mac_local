import { IsString, IsNotEmpty, IsOptional, IsBoolean, IsEnum } from 'class-validator';
import { ApiProperty } from '@nestjs/swagger';

export enum ShareLinkType {
  VIEW = 'view',
  EDIT = 'edit',
  EMBED = 'embed',
}

export enum ShareLinkScope {
  ANONYMOUS = 'anonymous',
  ORGANIZATION = 'organization',
}

export class ConvertWordToPdfDto {
  @ApiProperty({
    description: 'OneDrive file ID or sharing URL of the Word document',
    example: '01BYE5RZ6QN3ZWBTUFOFD3GSPGOHDJD36K',
  })
  @IsString()
  @IsNotEmpty()
  fileId: string;

  @ApiProperty({
    description: 'Whether to create a public shareable link for the PDF',
    example: true,
    required: false,
    default: true,
  })
  @IsBoolean()
  @IsOptional()
  createShareableLink?: boolean = true;

  @ApiProperty({
    description: 'Type of shareable link to create',
    enum: ShareLinkType,
    example: ShareLinkType.VIEW,
    required: false,
    default: ShareLinkType.VIEW,
  })
  @IsEnum(ShareLinkType)
  @IsOptional()
  shareLinkType?: ShareLinkType = ShareLinkType.VIEW;

  @ApiProperty({
    description: 'Scope of the shareable link',
    enum: ShareLinkScope,
    example: ShareLinkScope.ANONYMOUS,
    required: false,
    default: ShareLinkScope.ANONYMOUS,
  })
  @IsEnum(ShareLinkScope)
  @IsOptional()
  shareLinkScope?: ShareLinkScope = ShareLinkScope.ANONYMOUS;

  @ApiProperty({
    description: 'Custom filename for the PDF (without extension)',
    example: 'my-document',
    required: false,
  })
  @IsString()
  @IsOptional()
  customFilename?: string;
}

export class ConvertWordToPdfResponseDto {
  @ApiProperty({
    description: 'Success status of the conversion',
    example: true,
  })
  success: boolean;

  @ApiProperty({
    description: 'OneDrive file ID of the generated PDF',
    example: '01BYE5RZ6QN3ZWBTUFOFD3GSPGOHDJD36L',
  })
  pdfFileId: string;

  @ApiProperty({
    description: 'Direct download URL for the PDF file',
    example: 'https://graph.microsoft.com/v1.0/me/drive/items/01BYE5RZ6QN3ZWBTUFOFD3GSPGOHDJD36L/content',
  })
  downloadUrl: string;

  @ApiProperty({
    description: 'Public shareable link (if requested)',
    example: 'https://1drv.ms/b/s!AhKExample',
    required: false,
  })
  shareableUrl?: string;

  @ApiProperty({
    description: 'PDF file metadata',
  })
  metadata: {
    filename: string;
    size: number;
    createdDateTime: string;
    lastModifiedDateTime: string;
  };

  @ApiProperty({
    description: 'Any messages or warnings',
    required: false,
  })
  message?: string;
}

export class Office365AuthDto {
  @ApiProperty({
    description: 'Microsoft Azure AD Client ID',
    example: 'your-client-id',
  })
  @IsString()
  @IsNotEmpty()
  clientId: string;

  @ApiProperty({
    description: 'Microsoft Azure AD Client Secret',
    example: 'your-client-secret',
  })
  @IsString()
  @IsNotEmpty()
  clientSecret: string;

  @ApiProperty({
    description: 'Microsoft Azure AD Tenant ID',
    example: 'common',
  })
  @IsString()
  @IsNotEmpty()
  tenantId: string;

  @ApiProperty({
    description: 'Access token (if already authenticated)',
    required: false,
  })
  @IsString()
  @IsOptional()
  accessToken?: string;
}

export class BatchConvertWordToPdfDto {
  @ApiProperty({
    description: 'Array of OneDrive file IDs to convert',
    example: ['01BYE5RZ6QN3ZWBTUFOFD3GSPGOHDJD36K', '01BYE5RZ6QN3ZWBTUFOFD3GSPGOHDJD36M'],
  })
  @IsString({ each: true })
  @IsNotEmpty({ each: true })
  fileIds: string[];

  @ApiProperty({
    description: 'Whether to create public shareable links for all PDFs',
    example: true,
    required: false,
    default: true,
  })
  @IsBoolean()
  @IsOptional()
  createShareableLink?: boolean = true;

  @ApiProperty({
    description: 'Type of shareable link to create',
    enum: ShareLinkType,
    example: ShareLinkType.VIEW,
    required: false,
    default: ShareLinkType.VIEW,
  })
  @IsEnum(ShareLinkType)
  @IsOptional()
  shareLinkType?: ShareLinkType = ShareLinkType.VIEW;

  @ApiProperty({
    description: 'Scope of the shareable link',
    enum: ShareLinkScope,
    example: ShareLinkScope.ANONYMOUS,
    required: false,
    default: ShareLinkScope.ANONYMOUS,
  })
  @IsEnum(ShareLinkScope)
  @IsOptional()
  shareLinkScope?: ShareLinkScope = ShareLinkScope.ANONYMOUS;
}

export class BatchConvertWordToPdfResponseDto {
  @ApiProperty({
    description: 'Overall success status',
    example: true,
  })
  success: boolean;

  @ApiProperty({
    description: 'Number of successfully converted files',
    example: 2,
  })
  successCount: number;

  @ApiProperty({
    description: 'Number of failed conversions',
    example: 0,
  })
  failureCount: number;

  @ApiProperty({
    description: 'Conversion results for each file',
  })
  results: Array<{
    fileId: string;
    success: boolean;
    pdfFileId?: string;
    downloadUrl?: string;
    shareableUrl?: string;
    error?: string;
  }>;
}
