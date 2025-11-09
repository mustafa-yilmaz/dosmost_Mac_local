import {
  Controller,
  Get,
  Post,
  Put,
  Delete,
  Body,
  Query,
  Headers,
  UseGuards,
  HttpCode,
  HttpStatus,
} from '@nestjs/common';
import { GraphService } from './services/graph.service';
import {
  CreateRowDto,
  UpdateRowDto,
  DeleteRowDto,
  GetExcelDataDto,
} from './dto/excel.dto';
import { JwtAuthGuard } from '../../common/guards/jwt-auth.guard';

@Controller('api/office365')
@UseGuards(JwtAuthGuard)
export class Office365Controller {
  constructor(private readonly graphService: GraphService) {}

  /**
   * Get user's OneDrive information
   */
  @Get('drive')
  async getUserDrive(@Headers('x-ms-token') accessToken: string) {
    return this.graphService.getUserDrive(accessToken);
  }

  /**
   * List Excel files from user's OneDrive
   */
  @Get('excel/files')
  async listExcelFiles(
    @Headers('x-ms-token') accessToken: string,
    @Query('driveId') driveId?: string,
  ) {
    return this.graphService.listExcelFiles(accessToken, driveId);
  }

  /**
   * Get list of worksheets in a workbook
   */
  @Get('excel/worksheets')
  async getWorksheets(
    @Headers('x-ms-token') accessToken: string,
    @Query('driveId') driveId: string,
    @Query('itemId') itemId: string,
  ) {
    return this.graphService.getWorksheets(accessToken, driveId, itemId);
  }

  /**
   * Read data from Excel worksheet
   */
  @Get('excel/data')
  async getExcelData(
    @Headers('x-ms-token') accessToken: string,
    @Query() query: GetExcelDataDto,
  ) {
    return this.graphService.getExcelData(
      accessToken,
      query.driveId,
      query.itemId,
      query.worksheetName,
      query.range,
    );
  }

  /**
   * Create a new row in Excel worksheet
   */
  @Post('excel/row')
  @HttpCode(HttpStatus.CREATED)
  async createRow(
    @Headers('x-ms-token') accessToken: string,
    @Body() createRowDto: CreateRowDto,
  ) {
    return this.graphService.createRow(
      accessToken,
      createRowDto.driveId,
      createRowDto.itemId,
      createRowDto.worksheetName,
      createRowDto.data,
    );
  }

  /**
   * Update an existing row in Excel worksheet
   */
  @Put('excel/row')
  async updateRow(
    @Headers('x-ms-token') accessToken: string,
    @Body() updateRowDto: UpdateRowDto,
  ) {
    return this.graphService.updateRow(
      accessToken,
      updateRowDto.driveId,
      updateRowDto.itemId,
      updateRowDto.worksheetName,
      parseInt(updateRowDto.rowIndex),
      updateRowDto.data,
    );
  }

  /**
   * Delete a row from Excel worksheet
   */
  @Delete('excel/row')
  async deleteRow(
    @Headers('x-ms-token') accessToken: string,
    @Query() deleteRowDto: DeleteRowDto,
  ) {
    return this.graphService.deleteRow(
      accessToken,
      deleteRowDto.driveId,
      deleteRowDto.itemId,
      deleteRowDto.worksheetName,
      parseInt(deleteRowDto.rowIndex),
    );
  }
}
