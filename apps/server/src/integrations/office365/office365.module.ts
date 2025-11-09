import { Module } from '@nestjs/common';
import { Office365Controller } from './office365.controller';
import { GraphService } from './services/graph.service';

@Module({
  controllers: [Office365Controller],
  providers: [GraphService],
  exports: [GraphService],
})
export class Office365Module {}
