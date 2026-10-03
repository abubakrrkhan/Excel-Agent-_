import { Module } from '@nestjs/common';
import { BullModule } from '@nestjs/bullmq';
import { ConfigModule, ConfigService } from '@nestjs/config';
import { JobsController } from './jobs.controller';
import { JobsProcessor } from './jobs.processor';
import { AgentModule } from '../agent/agent.module';
import { ExcelModule } from '../excel/excel.module';
import { OperationsModule } from '../operations/operations.module';
import { FilesModule } from '../files/files.module';
import { EXCEL_AGENT_QUEUE } from './job.types';

@Module({
  imports: [
    // Register the queue (Redis connection from config)
    BullModule.registerQueueAsync({
      name: EXCEL_AGENT_QUEUE,
      imports: [ConfigModule],
      useFactory: (cfg: ConfigService) => ({
        connection: {
          // Railway provides REDIS_URL; fall back to individual host/port
          url: cfg.get<string>('app.valkey.url') || undefined,
          host: cfg.get<string>('app.valkey.host') || 'localhost',
          port: cfg.get<number>('app.valkey.port') || 6379,
        },
        defaultJobOptions: {
          attempts: 2,
          backoff: { type: 'exponential', delay: 2000 },
        },
      }),
      inject: [ConfigService],
    }),
    AgentModule,
    ExcelModule,
    OperationsModule,
    FilesModule,
  ],
  controllers: [JobsController],
  providers: [JobsProcessor],
  exports: [BullModule],
})
export class JobsModule {}
