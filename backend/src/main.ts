import { NestFactory } from '@nestjs/core';
import { ValidationPipe, Logger } from '@nestjs/common';
import { AppModule } from './app.module';

async function bootstrap() {
  const logger = new Logger('Bootstrap');
  const app = await NestFactory.create(AppModule);

  // Global prefix
  app.setGlobalPrefix('api');

  // CORS
  app.enableCors({
    origin: process.env.CORS_ORIGIN || 'http://localhost:3000',
    methods: ['GET', 'POST', 'PUT', 'PATCH', 'DELETE', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization'],
    credentials: true,
  });

  // Global validation pipe
  app.useGlobalPipes(
    new ValidationPipe({
      whitelist: true,
      forbidNonWhitelisted: true,
      transform: true,
    }),
  );

  // Production config warnings — catch misconfigured deployments early
  if (process.env.NODE_ENV === 'production') {
    if (!process.env.REDIS_URL && !process.env.VALKEY_URL) {
      logger.warn(
        'WARNING: Neither REDIS_URL nor VALKEY_URL is set. ' +
        'Falling back to redis://localhost:6379 — this will fail on Railway and most cloud hosts. ' +
        'Set REDIS_URL (Railway Redis plugin) or VALKEY_URL (self-hosted).',
      );
    }
    if (!process.env.AI_API_KEY) {
      logger.warn('WARNING: AI_API_KEY is not set — all agent requests will fail.');
    }
    if (!process.env.DATABASE_URL) {
      logger.warn('WARNING: DATABASE_URL is not set — database operations will fail.');
    }
    if (!process.env.CORS_ORIGIN) {
      logger.warn(
        'WARNING: CORS_ORIGIN is not set. Defaulting to http://localhost:3000 — ' +
        'cross-origin requests from your deployed frontend will be blocked.',
      );
    }
  }

  const port = process.env.PORT || 3001;
  await app.listen(port);

  logger.log(`Backend running on http://localhost:${port}/api`);
  logger.log(`Health check: http://localhost:${port}/api/health`);
}

bootstrap();

