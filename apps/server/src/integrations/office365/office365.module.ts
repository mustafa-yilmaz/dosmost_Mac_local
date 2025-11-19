import { Module } from '@nestjs/common';
import { ConfigModule } from '@nestjs/config';
import { Office365Controller } from './office365.controller';
import { GraphApiService } from './services/graph-api.service';
import { WordToPdfService } from './services/word-to-pdf.service';

@Module({
  imports: [ConfigModule],
  controllers: [Office365Controller],
  providers: [GraphApiService, WordToPdfService],
  exports: [GraphApiService, WordToPdfService],
})
export class Office365Module {}
