import { Module } from '@nestjs/common';
import { ScheduleModule } from '@nestjs/schedule';
import { BullModule } from '@nestjs/bullmq';
import { FilesService } from './files.service';
import { FilesController } from './files.controller';
import { CleanupService } from './cleanup.service';
import { ExcelModule } from '../excel/excel.module';
import { AiModule } from '../ai/ai.module';
import { OperationsModule } from '../operations/operations.module';
import { AgentModule } from '../agent/agent.module';
import { EXCEL_AGENT_QUEUE } from '../jobs/job.types';

@Module({
  imports: [
    // ScheduleModule provides SchedulerRegistry used by CleanupService for
    // dynamic cron registration (reads FILE_CLEANUP_CRON at startup).
    ScheduleModule.forRoot(),

    // Register queue reference so CleanupService can query active/waiting jobs
    // before cleanup — preventing deletion of files still in use.
    BullModule.registerQueue({ name: EXCEL_AGENT_QUEUE }),

    ExcelModule,
    AiModule,
    OperationsModule,
    AgentModule,
  ],
  controllers: [FilesController],
  providers: [FilesService, CleanupService],
  exports: [FilesService],
})
export class FilesModule {}
