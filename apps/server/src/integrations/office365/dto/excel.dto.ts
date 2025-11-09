import { IsString, IsNotEmpty, IsOptional, IsObject } from 'class-validator';

export class ExcelAuthDto {
  @IsString()
  @IsNotEmpty()
  accessToken: string;
}

export class CreateRowDto {
  @IsString()
  @IsNotEmpty()
  driveId: string;

  @IsString()
  @IsNotEmpty()
  itemId: string;

  @IsString()
  @IsNotEmpty()
  worksheetName: string;

  @IsObject()
  @IsNotEmpty()
  data: Record<string, any>;
}

export class UpdateRowDto {
  @IsString()
  @IsNotEmpty()
  driveId: string;

  @IsString()
  @IsNotEmpty()
  itemId: string;

  @IsString()
  @IsNotEmpty()
  worksheetName: string;

  @IsString()
  @IsNotEmpty()
  rowIndex: string;

  @IsObject()
  @IsNotEmpty()
  data: Record<string, any>;
}

export class DeleteRowDto {
  @IsString()
  @IsNotEmpty()
  driveId: string;

  @IsString()
  @IsNotEmpty()
  itemId: string;

  @IsString()
  @IsNotEmpty()
  worksheetName: string;

  @IsString()
  @IsNotEmpty()
  rowIndex: string;
}

export class GetExcelDataDto {
  @IsString()
  @IsNotEmpty()
  driveId: string;

  @IsString()
  @IsNotEmpty()
  itemId: string;

  @IsString()
  @IsNotEmpty()
  worksheetName: string;

  @IsOptional()
  @IsString()
  range?: string;
}
