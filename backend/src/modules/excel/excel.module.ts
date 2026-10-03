import { Module } from '@nestjs/common';
import { ExcelService } from './excel.service';
import { ExcelInspectionService } from './excel-inspection.service';
import { WorkbookStructureService } from './workbook-structure.service';

@Module({
  providers: [
    ExcelService,
    WorkbookStructureService,    // Step 2 — must be declared before ExcelInspectionService
    ExcelInspectionService,      // Step 1 — injects WorkbookStructureService
  ],
  exports: [
    ExcelService,
    ExcelInspectionService,
    WorkbookStructureService,
  ],
})
export class ExcelModule {}
