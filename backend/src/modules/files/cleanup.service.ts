/**
 * cleanup.service.ts
 *
 * Phase 12 — Scheduled file retention with active-job protection.
 *
 * Changes from Phase 11:
 *   1. Dynamic cron registration — reads FILE_CLEANUP_CRON at startup instead
 *      of using a hardcoded @Cron() decorator.
 *   2. Active-job protection — queries BullMQ for queued/active jobs before
 *      each cleanup pass and skips any files still referenced by those jobs.
 *
 * Configuration (environment variables):
 *   FILE_RETENTION_INPUT_HOURS   — hours to keep input files  (default: 24)
 *   FILE_RETENTION_OUTPUT_HOURS  — hours to keep output files (default: 2)
 *   FILE_CLEANUP_CRON            — cron expression (default: "0 3 * * *")
 */

import { Injectable, Logger, OnModuleInit, OnModuleDestroy } from '@nestjs/common';
import { SchedulerRegistry } from '@nestjs/schedule';
import { InjectQueue } from '@nestjs/bullmq';
import { Queue } from 'bullmq';
import { CronJob } from 'cron';
import { ConfigService } from '@nestjs/config';
import { FilesService } from './files.service';
import { EXCEL_AGENT_QUEUE } from '../jobs/job.types';

const CLEANUP_JOB_NAME = 'file-cleanup';

@Injectable()
export class CleanupService implements OnModuleInit, OnModuleDestroy {
  private readonly logger = new Logger(CleanupService.name);

  constructor(
    private readonly filesService: FilesService,
    private readonly configService: ConfigService,
    private readonly schedulerRegistry: SchedulerRegistry,
    @InjectQueue(EXCEL_AGENT_QUEUE) private readonly queue: Queue,
  ) {}

  // ── Lifecycle ─────────────────────────────────────────────────────────────

  onModuleInit() {
    const expression =
      this.configService.get<string>('app.cleanup.cron') ?? '0 3 * * *';

    let cronJob: CronJob;
    try {
      cronJob = new CronJob(expression, () => { void this.scheduledCleanup(); });
    } catch (err) {
      this.logger.error(
        `Invalid FILE_CLEANUP_CRON expression "${expression}": ${err.message}. ` +
          'Falling back to default "0 3 * * *".',
      );
      cronJob = new CronJob('0 3 * * *', () => { void this.scheduledCleanup(); });
    }

    // Remove stale registration if module is reloaded (e.g. hot-reload in dev)
    if (this.schedulerRegistry.doesExist('cron', CLEANUP_JOB_NAME)) {
      this.schedulerRegistry.deleteCronJob(CLEANUP_JOB_NAME);
    }

    this.schedulerRegistry.addCronJob(CLEANUP_JOB_NAME, cronJob as any);
    cronJob.start();

    this.logger.log(
      `File cleanup cron registered: "${expression}" (next: ${cronJob.nextDate().toISO()})`,
    );

    // Run an initial cleanup pass shortly after startup so stale files from a
    // previous run are removed immediately, without waiting for the first cron tick.
    // The 30-second delay gives the rest of the app (Redis, DB) time to fully
    // initialize before the cleanup attempts to query the job queue.
    setTimeout(() => {
      this.logger.log('Running initial startup cleanup pass...');
      void this.scheduledCleanup().catch(err =>
        this.logger.warn(`Startup cleanup failed (non-fatal): ${err.message}`),
      );
    }, 30_000);
  }

  onModuleDestroy() {
    try {
      if (this.schedulerRegistry.doesExist('cron', CLEANUP_JOB_NAME)) {
        const job = this.schedulerRegistry.getCronJob(CLEANUP_JOB_NAME);
        job.stop();
        this.schedulerRegistry.deleteCronJob(CLEANUP_JOB_NAME);
      }
    } catch {
      // Silent — module is shutting down
    }
  }

  // ── Scheduled cleanup ─────────────────────────────────────────────────────

  /**
   * Run the two-pass cleanup (outputs first, then inputs).
   *
   * Active-job protection: before deleting any file, this method queries
   * BullMQ for all waiting and active jobs and extracts their storedFileNames.
   * Those files are skipped even if they are past the retention age.
   *
   * This prevents the race condition where:
   *   1. A large file is uploaded and queued.
   *   2. Cleanup runs before the worker picks up the job.
   *   3. The worker cannot find the file.
   */
  async scheduledCleanup(): Promise<void> {
    const inputMs =
      this.configService.get<number>('app.cleanup.inputRetentionMs') ??
      24 * 60 * 60 * 1000;
    const outputMs =
      this.configService.get<number>('app.cleanup.outputRetentionMs') ??
      2 * 60 * 60 * 1000;

    this.logger.log(
      `Cleanup starting | input: ${inputMs / 3600000}h | output: ${outputMs / 3600000}h`,
    );

    // Build the protected-files set ONCE before both passes
    const protectedFiles = await this.getActiveJobFiles();

    try {
      const outputResult = await this.filesService.cleanupOldFiles(
        outputMs, 'output', protectedFiles,
      );
      this.logger.log(
        `Output cleanup: ${outputResult.deleted} deleted, ` +
          `${outputResult.errors} errors`,
      );

      const inputResult = await this.filesService.cleanupOldFiles(
        inputMs, 'input', protectedFiles,
      );
      this.logger.log(
        `Input cleanup: ${inputResult.deleted} deleted, ` +
          `${inputResult.errors} errors`,
      );
    } catch (err) {
      this.logger.error(`Cleanup failed: ${err.message}`);
    }
  }

  /**
   * Query BullMQ for all waiting and active jobs and return the set of
   * stored filenames they reference.
   *
   * These files must not be deleted even if they are past the retention age.
   */
  private async getActiveJobFiles(): Promise<Set<string>> {
    const protected_ = new Set<string>();
    try {
      // Guard all states that still need the uploaded file:
      //   waiting  — queued, not yet started
      //   active   — currently being processed
      //   delayed  — scheduled to run later, OR waiting between retry attempts
      //              (BullMQ moves failed-but-retrying jobs to 'delayed')
      //   failed   — exhausted retries but file may still be referenced
      const [waiting, active, delayed, failed] = await Promise.all([
        this.queue.getWaiting(0, -1),
        this.queue.getActive(0, -1),
        this.queue.getDelayed(0, -1),
        this.queue.getFailed(0, -1),
      ]);
      for (const job of [...waiting, ...active, ...delayed, ...failed]) {
        if (job.data?.storedFileName) {
          protected_.add(job.data.storedFileName);
        }
      }
      if (protected_.size > 0) {
        this.logger.log(
          `Cleanup: ${protected_.size} file(s) protected (waiting/active/delayed/failed jobs)`,
        );
      }
    } catch (err) {
      // If BullMQ is temporarily unavailable, err on the side of caution:
      // abort the cleanup rather than potentially deleting active job files.
      this.logger.error(
        `Cleanup: could not query job queue — aborting to protect active files: ${err.message}`,
      );
      throw err;
    }
    return protected_;
  }

  // ── Manual trigger ────────────────────────────────────────────────────────

  /**
   * Manual trigger for testing and admin operations.
   * Accepts explicit retention overrides; still respects active-job protection.
   */
  async runCleanup(inputMs?: number, outputMs?: number): Promise<{
    input: { deleted: number; errors: number };
    output: { deleted: number; errors: number };
  }> {
    const effectiveInputMs =
      inputMs ??
      this.configService.get<number>('app.cleanup.inputRetentionMs') ??
      24 * 60 * 60 * 1000;
    const effectiveOutputMs =
      outputMs ??
      this.configService.get<number>('app.cleanup.outputRetentionMs') ??
      2 * 60 * 60 * 1000;

    const protectedFiles = await this.getActiveJobFiles();

    const output = await this.filesService.cleanupOldFiles(
      effectiveOutputMs, 'output', protectedFiles,
    );
    const input = await this.filesService.cleanupOldFiles(
      effectiveInputMs, 'input', protectedFiles,
    );
    return { input, output };
  }
}
