import { Injectable, OnModuleInit, OnModuleDestroy, Logger } from '@nestjs/common';
import { Pool } from 'pg';

/**
 * DatabaseService wraps the pg Pool for local dev / type-check.
 * In production, swap this body for PrismaClient extends once
 * `prisma generate` runs on the server (where binaries.prisma.sh is reachable).
 *
 * The Prisma schema at prisma/schema.prisma is production-authoritative.
 */
@Injectable()
export class DatabaseService implements OnModuleInit, OnModuleDestroy {
  private readonly logger = new Logger(DatabaseService.name);
  private pool: Pool;

  async onModuleInit() {
    const databaseUrl = process.env.DATABASE_URL;
    if (!databaseUrl) {
      this.logger.warn('DATABASE_URL not set — skipping DB connection');
      return;
    }

    try {
      this.pool = new Pool({ connectionString: databaseUrl });
      await this.pool.query('SELECT 1');
      this.logger.log('✅ Database connected');
    } catch (error) {
      this.logger.warn(`Database not available: ${error.message} (will retry when DB is ready)`);
    }
  }

  async onModuleDestroy() {
    if (this.pool) {
      await this.pool.end();
      this.logger.log('Database disconnected');
    }
  }

  getPool(): Pool {
    return this.pool;
  }
}
