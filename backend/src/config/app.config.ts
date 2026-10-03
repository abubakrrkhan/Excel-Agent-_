import { registerAs } from '@nestjs/config';

export default registerAs('app', () => ({
  port: parseInt(process.env.PORT, 10) || 3001,
  nodeEnv: process.env.NODE_ENV || 'development',
  corsOrigin: process.env.CORS_ORIGIN || 'http://localhost:3000',

  database: {
    url: process.env.DATABASE_URL,
  },

  valkey: {
    host: process.env.VALKEY_HOST || 'localhost',
    port: parseInt(process.env.VALKEY_PORT, 10) || 6379,
    /**
     * Redis/Valkey connection URL.
     *
     * Precedence (highest to lowest):
     *   1. REDIS_URL   — set automatically by Railway's Redis plugin
     *   2. VALKEY_URL  — used for self-hosted Valkey / other providers
     *   3. Default     — redis://localhost:6379 (local development)
     *
     * Both variables are accepted so the same codebase works on Railway
     * (which sets REDIS_URL) and on Docker Compose / bare-metal (VALKEY_URL).
     * Never log the resolved URL — it may contain credentials.
     */
    url: process.env.REDIS_URL || process.env.VALKEY_URL || 'redis://localhost:6379',
  },

  ai: {
    apiKey: process.env.AI_API_KEY,
    model: process.env.AI_MODEL || 'claude-sonnet-4-6',
    baseUrl: process.env.AI_BASE_URL || 'https://api.anthropic.com',
  },

  cleanup: {
    inputRetentionMs:
      parseInt(process.env.FILE_RETENTION_INPUT_HOURS || '24', 10) * 60 * 60 * 1000,
    outputRetentionMs:
      parseInt(process.env.FILE_RETENTION_OUTPUT_HOURS || '2', 10) * 60 * 60 * 1000,
    /**
     * Cron expression controlling the cleanup schedule.
     * Default: "0 3 * * *" (daily at 03:00).
     * Override with FILE_CLEANUP_CRON env var.
     * The CleanupService reads this at startup and registers a dynamic job.
     */
    cron: process.env.FILE_CLEANUP_CRON || '0 3 * * *',
  },

  rateLimit: {
    /** Maximum requests per TTL window per IP on rate-limited endpoints. */
    max: parseInt(process.env.RATE_LIMIT_MAX || '20', 10),
    /** TTL window in seconds. */
    ttlSeconds: parseInt(process.env.RATE_LIMIT_TTL || '60', 10),
  },
}));
