import { Module } from '@nestjs/common';
import { APP_GUARD } from '@nestjs/core';
import { ConfigModule, ConfigService } from '@nestjs/config';
import { BullModule } from '@nestjs/bullmq';
import { ThrottlerModule, ThrottlerGuard } from '@nestjs/throttler';
import { HealthModule } from './modules/health/health.module';
import { DatabaseModule } from './modules/database/database.module';
import { ExcelModule } from './modules/excel/excel.module';
import { AiModule } from './modules/ai/ai.module';
import { FilesModule } from './modules/files/files.module';
import { OperationsModule } from './modules/operations/operations.module';
import { AgentModule } from './modules/agent/agent.module';
import { JobsModule } from './modules/jobs/jobs.module';
import appConfig from './config/app.config';

@Module({
  imports: [
    ConfigModule.forRoot({
      isGlobal: true,
      load: [appConfig],
      envFilePath: ['.env'],
    }),

    // Rate limiting — global guard is registered in `providers` below.
    // Limits are controlled by RATE_LIMIT_MAX (default: 20) and
    // RATE_LIMIT_TTL (default: 60 s). Endpoints with @SkipThrottle() are exempt.
    ThrottlerModule.forRootAsync({
      imports: [ConfigModule],
      useFactory: (cfg: ConfigService) => ([{
        ttl: (cfg.get<number>('app.rateLimit.ttlSeconds') ?? 60) * 1000,
        limit: cfg.get<number>('app.rateLimit.max') ?? 20,
      }]),
      inject: [ConfigService],
    }),

    // BullMQ — global Redis connection.
    // URL resolution order: REDIS_URL (Railway) → VALKEY_URL → localhost default.
    // retryStrategy: exponential back-off up to 10 s, stops after 10 attempts.
    BullModule.forRootAsync({
      imports: [ConfigModule],
      useFactory: (cfg: ConfigService) => ({
        connection: {
          url: cfg.get<string>('app.valkey.url') || undefined,
          host: cfg.get<string>('app.valkey.host') || 'localhost',
          port: cfg.get<number>('app.valkey.port') || 6379,
          retryStrategy: (attempt: number) => {
            if (attempt > 10) return null;
            return Math.min(attempt * 500, 10_000);
          },
          maxRetriesPerRequest: null,
          enableReadyCheck: false,
        },
      }),
      inject: [ConfigService],
    }),

    DatabaseModule,
    HealthModule,
    ExcelModule,
    AiModule,
    FilesModule,
    OperationsModule,
    AgentModule,
    JobsModule,
  ],

  providers: [
    // Apply ThrottlerGuard globally. Endpoints opt out with @SkipThrottle().
    // Without this provider, @Throttle() decorators are metadata only — no effect.
    {
      provide: APP_GUARD,
      useClass: ThrottlerGuard,
    },
  ],
})
export class AppModule {}
