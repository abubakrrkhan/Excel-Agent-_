/**
 * jobs.controller.ts
 *
 * Phase 6 — Job submission and status polling endpoints.
 *
 * POST /api/jobs/submit  — submit file + instruction → returns { jobId }
 * GET  /api/jobs/:id     — poll job status → returns AgentJobResult
 */

import {
  Controller, Post, Get, Param, Body,
  UploadedFiles, UseInterceptors,
  HttpCode, HttpStatus, Logger, BadRequestException,
} from '@nestjs/common';
import { FileFieldsInterceptor } from '@nestjs/platform-express';
import { memoryStorage } from 'multer';
import * as crypto from 'crypto';
import { InjectQueue } from '@nestjs/bullmq';
import { Queue } from 'bullmq';
import { SkipThrottle } from '@nestjs/throttler';
import {
  EXCEL_AGENT_QUEUE,
  EXCEL_AGENT_JOB,
  AgentJobResult,
} from './job.types';
import { FilesService } from '../files/files.service';

const MAX_FILE_SIZE = 10 * 1024 * 1024; // 10 MB — same limit as FilesController
const MULTER_OPTIONS = {
  storage: memoryStorage(),
  limits: { fileSize: MAX_FILE_SIZE },
};

interface SubmitResponse {
  jobId: string;
  status: 'queued';
  message: string;
  pollUrl: string;
}

@Controller('jobs')
export class JobsController {
  private readonly logger = new Logger(JobsController.name);

  constructor(
    @InjectQueue(EXCEL_AGENT_QUEUE) private readonly queue: Queue,
    private readonly filesService: FilesService,
  ) {}

  /**
   * POST /api/jobs/submit
   *
   * Rate-limited by the global ThrottlerGuard (RATE_LIMIT_MAX requests per
   * RATE_LIMIT_TTL seconds per IP). Returns HTTP 429 when exceeded.
   * No per-route @Throttle() override — the module-level config governs this.
   */
  @Post('submit')
  @HttpCode(HttpStatus.ACCEPTED)
  @UseInterceptors(
    FileFieldsInterceptor([{ name: 'file', maxCount: 1 }], MULTER_OPTIONS),
  )
  async submitJob(
    @UploadedFiles() files: { file?: Express.Multer.File[] },
    @Body('instruction') instruction: string,
    @Body('clarificationAnswer') clarificationAnswer?: string,
  ): Promise<SubmitResponse> {
    const file = files?.file?.[0];
    this.filesService.validateExcelFile(file);

    if (!instruction?.trim()) {
      throw new BadRequestException('Instruction is required');
    }

    const jobId = crypto.randomUUID();

    // Save file to disk — NEVER embed file bytes in the queue payload.
    // A 10 MB file ≈ 13 MB base64; storing that in Redis exhausts memory fast.
    const storedFileName = await this.filesService.saveFileToDisk(file);

    await this.queue.add(
      EXCEL_AGENT_JOB,
      {
        jobId,
        storedFileName,
        originalName: file.originalname,
        instruction,
        clarificationAnswer: clarificationAnswer?.trim() || undefined,
      },
      {
        jobId,
        attempts: 2,
        backoff: { type: 'exponential', delay: 3000 },
        removeOnComplete: { age: 3600 },  // keep results 1 hour
        removeOnFail: { age: 86400 },     // keep failures 24 hours
      },
    );

    this.logger.log(`Job queued: ${jobId} | file=${storedFileName} | "${instruction.slice(0, 60)}"`);

    return {
      jobId,
      status: 'queued',
      message: 'Job submitted. Poll /api/jobs/:id for status.',
      pollUrl: `/api/jobs/${jobId}`,
    };
  }

  /**
   * GET /api/jobs/:id
   * Poll job status. Not rate-limited — frontend calls every 2 s.
   */
  @SkipThrottle()
  @Get(':id')
  async getJobStatus(@Param('id') jobId: string): Promise<AgentJobResult> {
    const job = await this.queue.getJob(jobId);

    if (!job) {
      return {
        status: 'failed',
        jobId,
        error: 'Job not found. It may have expired or the ID is incorrect.',
      };
    }

    const state = await job.getState();

    // Job is still in queue
    if (state === 'waiting' || state === 'delayed') {
      return { status: 'queued', jobId };
    }

    // Job is actively being processed
    if (state === 'active') {
      const progress = typeof job.progress === 'number' ? job.progress : 0;
      const result: AgentJobResult = { status: 'processing', jobId };
      // Return partial activity log if available (set by processor)
      if (job.data) {
        (result as any).progress = progress;
      }
      return result;
    }

    // Job completed — return processor result
    if (state === 'completed' && job.returnvalue) {
      return job.returnvalue as AgentJobResult;
    }

    // Job failed
    if (state === 'failed') {
      return {
        status: 'failed',
        jobId,
        error: job.failedReason ?? 'Job processing failed',
        finishedAt: new Date(job.finishedOn ?? Date.now()).toISOString(),
      };
    }

    return { status: 'processing', jobId };
  }
}
