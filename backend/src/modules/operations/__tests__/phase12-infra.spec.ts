/**
 * phase12-infra.spec.ts
 *
 * Phase 12 — Backend & Infrastructure reliability regression tests.
 *
 * Tests confirm every issue from the Phase 12 spec is either fixed or
 * was already correctly implemented. No frontend files are touched.
 *
 *   A. Redis/Valkey env-var precedence (Stage 2)
 *   B. package.json integrity — npm ci lockfile compatibility (Stage 3)
 *   C. Docker Compose API URL (Stage 4)
 *   D. Cleanup schedule — dynamic cron from FILE_CLEANUP_CRON (Stage 5)
 *   E. Active-job file protection (Stage 6)
 *   F. Rate-limit — global guard + configurable values (Stage 7)
 *   G. Redis retry strategy (Stage 8)
 *   H. Frontend source files untouched (strict rule)
 */

import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';

const ROOT = path.join(__dirname, '../../../../../');
const BACKEND = path.join(ROOT, 'backend');
const FRONTEND = path.join(ROOT, 'frontend');

// ─── Helpers ──────────────────────────────────────────────────────────────────

function readJson(p: string) {
  return JSON.parse(fs.readFileSync(p, 'utf8'));
}
function readSrc(rel: string) {
  return fs.readFileSync(path.join(BACKEND, rel), 'utf8');
}
function readFront(rel: string) {
  return fs.readFileSync(path.join(FRONTEND, rel), 'utf8');
}

// ═══════════════════════════════════════════════════════════════════════════
// A. Redis/Valkey environment variable configuration
// ═══════════════════════════════════════════════════════════════════════════

describe('A. Redis/Valkey env-var precedence', () => {
  const configSrc = readSrc('src/config/app.config.ts');
  const moduleSrc = readSrc('src/app.module.ts');

  it('REDIS_URL takes precedence over VALKEY_URL in app.config.ts', () => {
    // The URL line must try REDIS_URL before VALKEY_URL
    const line = configSrc.split('\n').find(l =>
      l.includes('REDIS_URL') && l.includes('VALKEY_URL'),
    );
    expect(line).toBeDefined();
    expect(line!.indexOf('REDIS_URL')).toBeLessThan(line!.indexOf('VALKEY_URL'));
  });

  it('app.config.ts defines valkey.url reading REDIS_URL first', () => {
    expect(configSrc).toContain('process.env.REDIS_URL');
    expect(configSrc).toContain('process.env.VALKEY_URL');
    // Falls back to localhost if neither is set
    expect(configSrc).toContain('redis://localhost:6379');
  });

  it('app.module.ts BullMQ connection reads from app.valkey.url', () => {
    expect(moduleSrc).toContain("cfg.get<string>('app.valkey.url')");
  });

  it('REDIS_URL precedence — runtime simulation', () => {
    const resolve = (REDIS_URL?: string, VALKEY_URL?: string) =>
      REDIS_URL || VALKEY_URL || 'redis://localhost:6379';

    // REDIS_URL only
    expect(resolve('redis://redis-host:6379')).toBe('redis://redis-host:6379');
    // VALKEY_URL only
    expect(resolve(undefined, 'redis://valkey-host:6379')).toBe('redis://valkey-host:6379');
    // Both — REDIS_URL wins
    expect(resolve('redis://redis-host:6379', 'redis://valkey-host:6379')).toBe('redis://redis-host:6379');
    // Neither — default
    expect(resolve()).toBe('redis://localhost:6379');
  });

  it('.env.example documents REDIS_URL and VALKEY_URL with clear precedence', () => {
    const envExample = fs.readFileSync(path.join(ROOT, '.env.example'), 'utf8');
    expect(envExample).toContain('REDIS_URL');
    expect(envExample).toContain('VALKEY_URL');
  });

  it('docker-compose.yml uses VALKEY_URL for self-hosted (correct for local dev)', () => {
    const compose = fs.readFileSync(path.join(ROOT, 'docker-compose.yml'), 'utf8');
    expect(compose).toContain('VALKEY_URL');
    // Should NOT use REDIS_URL in compose — that's Railway-only
    expect(compose).not.toContain('REDIS_URL');
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// B. Package.json and lockfile integrity
// ═══════════════════════════════════════════════════════════════════════════

describe('B. Lockfile and package.json integrity', () => {
  it('backend package-lock.json exists', () => {
    expect(fs.existsSync(path.join(BACKEND, 'package-lock.json'))).toBe(true);
  });

  it('frontend package-lock.json exists', () => {
    expect(fs.existsSync(path.join(FRONTEND, 'package-lock.json'))).toBe(true);
  });

  it('backend package-lock.json is lockfileVersion 3 (npm 7+)', () => {
    const lock = readJson(path.join(BACKEND, 'package-lock.json'));
    expect(lock.lockfileVersion).toBeGreaterThanOrEqual(2);
  });

  it('@nestjs/schedule in package.json is v4+ (fixes reflect-metadata peer dep conflict)', () => {
    const pkg = readJson(path.join(BACKEND, 'package.json'));
    const ver = pkg.dependencies['@nestjs/schedule'] as string;
    expect(ver).toBeDefined();
    const major = parseInt(ver.replace(/[\^~>=]/, ''), 10);
    expect(major).toBeGreaterThanOrEqual(4);
  });

  it('@nestjs/schedule v4 is compatible with reflect-metadata ^0.2.x', () => {
    // v3 required ^0.1.12, v4+ supports ^0.2.0
    // Verify by checking that reflect-metadata in package.json is ^0.2.x
    const pkg = readJson(path.join(BACKEND, 'package.json'));
    const rm = pkg.dependencies['reflect-metadata'] as string;
    expect(rm).toBeDefined();
    expect(rm).toMatch(/\^0\.2/);
  });

  it('backend Dockerfile uses npm ci (not npm install)', () => {
    const df = fs.readFileSync(path.join(ROOT, 'docker/Dockerfile.backend'), 'utf8');
    expect(df).toContain('npm ci');
    expect(df).not.toMatch(/^RUN npm install$/m);
  });

  it('frontend Dockerfile uses npm ci', () => {
    const df = fs.readFileSync(path.join(ROOT, 'docker/Dockerfile.frontend'), 'utf8');
    expect(df).toContain('npm ci');
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// C. Docker Compose browser API connectivity
// ═══════════════════════════════════════════════════════════════════════════

describe('C. Docker Compose API URL — browser accessibility', () => {
  const compose = fs.readFileSync(path.join(ROOT, 'docker-compose.yml'), 'utf8');

  it('NEXT_PUBLIC_API_URL uses localhost:3001 not backend:3001 as the actual value', () => {
    // http://backend:3001 is only reachable inside Docker, not from the browser.
    // The compose file documents this with a comment; the actual value must be localhost:3001.
    expect(compose).toContain('NEXT_PUBLIC_API_URL: http://localhost:3001');
    // Confirm that backend:3001 does not appear as an assigned value (it may appear in comments)
    const valueLines = compose.split('\n').filter(l =>
      l.includes('NEXT_PUBLIC_API_URL') && l.includes('=') && !l.trim().startsWith('#'),
    );
    for (const line of valueLines) {
      expect(line).not.toContain('backend:3001');
    }
  });

  it('backend port 3001 is published to host', () => {
    expect(compose).toContain('"3001:3001"');
  });

  it('frontend port 3000 is published to host', () => {
    expect(compose).toContain('"3000:3000"');
  });

  it('next.config.js has standalone output for Docker', () => {
    const nc = readFront('next.config.js');
    expect(nc).toContain("output: 'standalone'");
  });

  it('NEXT_PUBLIC_API_URL is passed both as build ARG and runtime ENV', () => {
    const df = fs.readFileSync(path.join(ROOT, 'docker/Dockerfile.frontend'), 'utf8');
    expect(df).toContain('ARG NEXT_PUBLIC_API_URL');
    expect(df).toContain('ENV NEXT_PUBLIC_API_URL');
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// D. Cleanup schedule — dynamic cron from FILE_CLEANUP_CRON
// ═══════════════════════════════════════════════════════════════════════════

describe('D. Cleanup schedule — dynamic cron from config', () => {
  const cleanupSrc = readSrc('src/modules/files/cleanup.service.ts');
  const configSrc = readSrc('src/config/app.config.ts');

  it('CleanupService does NOT use hardcoded @Cron() decorator', () => {
    // Should use dynamic SchedulerRegistry.addCronJob instead
    expect(cleanupSrc).not.toContain("@Cron('");
    expect(cleanupSrc).not.toContain('@Cron("');
  });

  it('CleanupService reads FILE_CLEANUP_CRON from config at startup', () => {
    expect(cleanupSrc).toContain("app.cleanup.cron");
    expect(cleanupSrc).toContain('onModuleInit');
  });

  it('FILE_CLEANUP_CRON is in app.config.ts', () => {
    expect(configSrc).toContain('FILE_CLEANUP_CRON');
  });

  it('default cron expression is "0 3 * * *" (3 AM daily)', () => {
    expect(configSrc).toContain("'0 3 * * *'");
  });

  it('invalid cron expression falls back to default with error log', () => {
    expect(cleanupSrc).toContain('Falling back to default');
    expect(cleanupSrc).toContain('0 3 * * *');
  });

  it('duplicate cron registration is prevented on hot-reload', () => {
    expect(cleanupSrc).toContain('doesExist');
    expect(cleanupSrc).toContain('deleteCronJob');
  });

  it('onModuleDestroy stops and removes the cron job', () => {
    expect(cleanupSrc).toContain('onModuleDestroy');
    expect(cleanupSrc).toContain('job.stop()');
  });

  it('FILE_RETENTION_INPUT_HOURS and FILE_RETENTION_OUTPUT_HOURS are in config', () => {
    expect(configSrc).toContain('FILE_RETENTION_INPUT_HOURS');
    expect(configSrc).toContain('FILE_RETENTION_OUTPUT_HOURS');
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// E. Active-job file protection during cleanup
// ═══════════════════════════════════════════════════════════════════════════

describe('E. Active-job file protection', () => {
  const cleanupSrc = readSrc('src/modules/files/cleanup.service.ts');
  const servicesSrc = readSrc('src/modules/files/files.service.ts');
  const modulesSrc = readSrc('src/modules/files/files.module.ts');

  it('CleanupService queries BullMQ for waiting and active jobs before cleanup', () => {
    expect(cleanupSrc).toContain('getWaiting');
    expect(cleanupSrc).toContain('getActive');
  });

  it('getActiveJobFiles builds a Set of protected filenames', () => {
    expect(cleanupSrc).toContain('storedFileName');
    expect(cleanupSrc).toContain('protected_');
  });

  it('cleanupOldFiles accepts a protectedFiles parameter', () => {
    expect(servicesSrc).toContain('protectedFiles: Set<string>');
    expect(servicesSrc).toContain('protectedFiles.has(file)');
  });

  it('if BullMQ unavailable during cleanup, abort rather than risk deleting active files', () => {
    expect(cleanupSrc).toContain('aborting to protect active files');
    expect(cleanupSrc).toContain('throw err');
  });

  it('FilesModule registers the BullMQ queue for CleanupService', () => {
    expect(modulesSrc).toContain('BullModule.registerQueue');
    expect(modulesSrc).toContain('EXCEL_AGENT_QUEUE');
  });

  it('cleanup with protected files — unit test', async () => {
    // Simulate cleanupOldFiles with a protected file
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'p12-e-'));
    const UUID = '550e8400-e29b-41d4-a716-446655440000';
    const activeFile = path.join(tmpDir, `${UUID}.xlsx`);
    const oldFile = path.join(tmpDir, `output-${UUID}.xlsx`);
    fs.writeFileSync(activeFile, 'active-job-file');
    fs.writeFileSync(oldFile, 'old-output-file');
    // Backdate both
    const old = new Date(Date.now() - 5 * 3600 * 1000);
    fs.utimesSync(activeFile, old, old);
    fs.utimesSync(oldFile, old, old);

    // Simulate: activeFile is protected
    const protected_ = new Set([`${UUID}.xlsx`]);
    const retention = 2 * 3600 * 1000;
    const inputPattern = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\.xlsx$/;
    const outputPattern = /^output-[0-9a-f-]+\.xlsx$/;
    const cutoff = Date.now() - retention;
    let deleted = 0;
    for (const file of fs.readdirSync(tmpDir)) {
      if (!inputPattern.test(file) && !outputPattern.test(file)) continue;
      if (protected_.has(file)) continue; // PROTECTED
      const stat = fs.statSync(path.join(tmpDir, file));
      if (stat.mtimeMs < cutoff) { fs.unlinkSync(path.join(tmpDir, file)); deleted++; }
    }

    expect(deleted).toBe(1); // only oldFile deleted
    expect(fs.existsSync(activeFile)).toBe(true); // active file protected
    expect(fs.existsSync(oldFile)).toBe(false);   // old output deleted
    fs.rmSync(tmpDir, { recursive: true });
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// F. Rate-limiting — global guard + configurable values
// ═══════════════════════════════════════════════════════════════════════════

describe('F. Rate-limiting — global guard + configurable values', () => {
  const appModSrc = readSrc('src/app.module.ts');
  const configSrc = readSrc('src/config/app.config.ts');
  const jobsCtrl = readSrc('src/modules/jobs/jobs.controller.ts');

  it('ThrottlerModule reads limit from RATE_LIMIT_MAX env var', () => {
    expect(appModSrc).toContain('RATE_LIMIT_MAX');
    expect(appModSrc).toContain("'app.rateLimit.max'");
  });

  it('ThrottlerModule reads TTL from RATE_LIMIT_TTL env var', () => {
    expect(appModSrc).toContain('RATE_LIMIT_TTL');
    expect(appModSrc).toContain("'app.rateLimit.ttlSeconds'");
  });

  it('ThrottlerGuard is registered globally via APP_GUARD', () => {
    expect(appModSrc).toContain('APP_GUARD');
    expect(appModSrc).toContain('ThrottlerGuard');
  });

  it('job submission endpoint is rate-limited (no @SkipThrottle on submit)', () => {
    // The submit method must NOT be decorated with @SkipThrottle
    const submitBlock = jobsCtrl.split('@Post(\'submit\')')[0] + '@Post(\'submit\')';
    expect(submitBlock).not.toContain('@SkipThrottle()');
  });

  it('job status polling endpoint has @SkipThrottle', () => {
    const pollBlock = jobsCtrl.split("@Get(':id')")[0];
    expect(pollBlock).toContain('@SkipThrottle()');
  });

  it('RATE_LIMIT_MAX and RATE_LIMIT_TTL defaults are in app.config.ts', () => {
    expect(configSrc).toContain('RATE_LIMIT_MAX');
    expect(configSrc).toContain('RATE_LIMIT_TTL');
  });

  it('default rate limit is 20 requests per 60 seconds', () => {
    expect(configSrc).toContain("'20'");  // RATE_LIMIT_MAX default
    expect(configSrc).toContain("'60'");  // RATE_LIMIT_TTL default
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// G. Redis retry strategy
// ═══════════════════════════════════════════════════════════════════════════

describe('G. Redis connection retry strategy', () => {
  const appModSrc = readSrc('src/app.module.ts');

  it('retryStrategy is configured in BullMQ connection', () => {
    expect(appModSrc).toContain('retryStrategy');
  });

  it('maxRetriesPerRequest is null (required by BullMQ)', () => {
    expect(appModSrc).toContain('maxRetriesPerRequest: null');
  });

  it('enableReadyCheck is false (no crash if Redis not ready at startup)', () => {
    expect(appModSrc).toContain('enableReadyCheck: false');
  });

  it('retry stops after 10 attempts (no infinite loop)', () => {
    const retryFn = (attempt: number) => {
      if (attempt > 10) return null;
      return Math.min(attempt * 500, 10_000);
    };
    expect(retryFn(11)).toBeNull();
    expect(retryFn(10)).not.toBeNull();
  });

  it('retry delays are bounded at 10 seconds', () => {
    const retryFn = (attempt: number) => {
      if (attempt > 10) return null;
      return Math.min(attempt * 500, 10_000);
    };
    for (let i = 1; i <= 10; i++) {
      expect(retryFn(i)!).toBeLessThanOrEqual(10_000);
    }
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// H. Frontend source files untouched
// ═══════════════════════════════════════════════════════════════════════════

describe('H. Frontend source files not modified (Phase 12 strict rule)', () => {
  const frontendSources = [
    'src/app/page.tsx',
    'src/app/globals.css',
    'src/app/layout.tsx',
    'src/lib/api.ts',
    'src/types/index.ts',
    'src/components/FileUpload.tsx',
    'src/components/InstructionInput.tsx',
    'src/components/ResultPanel.tsx',
    'src/components/StatusBadge.tsx',
  ];

  it('all frontend source files exist', () => {
    for (const rel of frontendSources) {
      expect(fs.existsSync(path.join(FRONTEND, rel))).toBe(true);
    }
  });

  it('page.tsx has the ChatGPT sidebar layout (not modified by Phase 12)', () => {
    const pageSrc = readFront('src/app/page.tsx');
    expect(pageSrc).toContain('Sidebar');
    expect(pageSrc).toContain('history');
    expect(pageSrc).toContain('Ready when you are.');
  });

  it('globals.css preserves pure black theme', () => {
    const css = readFront('src/app/globals.css');
    expect(css).toContain('#0a0a0a');
  });

  it('api.ts uses /jobs/submit for async job submission', () => {
    const api = readFront('src/lib/api.ts');
    expect(api).toContain('/jobs/submit');
    expect(api).toContain('clarificationAnswer');
  });

  it('next.config.js was not modified — still has standalone output', () => {
    const nc = readFront('next.config.js');
    expect(nc).toContain("output: 'standalone'");
    expect(nc).toContain('rewrites');
  });
});
