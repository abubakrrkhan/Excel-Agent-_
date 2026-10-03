import { Module } from '@nestjs/common';
import { OperationsService } from './operations.service';
import { OperationExecutionRouter } from './operation-execution-router';
import { ExcelModule } from '../excel/excel.module';

@Module({
  imports: [
    ExcelModule, // provides ExcelInspectionService + WorkbookStructureService + ExcelService
  ],
  providers: [OperationsService, OperationExecutionRouter],
  exports: [OperationsService, OperationExecutionRouter],
})
export class OperationsModule {}
