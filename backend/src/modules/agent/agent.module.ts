import { Module } from '@nestjs/common';
import { AgentService } from './agent.service';
import { AiModule } from '../ai/ai.module';
import { ExcelModule } from '../excel/excel.module';
import { OperationsModule } from '../operations/operations.module';

@Module({
  imports: [
    AiModule,          // AiService.callAgentRound
    ExcelModule,       // ExcelInspectionService + WorkbookStructureService
    OperationsModule,  // OperationsService.validatePlan + executeSetCellByLabel
  ],
  providers: [AgentService],
  exports: [AgentService],
})
export class AgentModule {}
