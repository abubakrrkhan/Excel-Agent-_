/**
 * jobs.processor.ts
 *
 * Phase 6 — BullMQ job processor.
 *
 * Runs the full agent pipeline in a background worker:
 *   1. Decode workbook buffer from job payload
 *   2. Run AgentService.runAgent() (multi-round tool-calling)
 *   3. If completed: validate plan, apply ops, save output
 *   4. Return structured AgentJobResult
 *
 * This processor runs outside the HTTP request lifecycle,
 * so there's no timeout risk for long-running agent sessions.
 */

import { Processor, WorkerHost } from '@nestjs/bullmq';
import { Logger } from '@nestjs/common';
import { Job } from 'bullmq';
import { AgentService } from '../agent/agent.service';
import { ExcelService } from '../excel/excel.service';
import { OperationsService } from '../operations/operations.service';
import { OperationExecutionRouter } from '../operations/operation-execution-router';
import { FilesService } from '../files/files.service';
import {
  EXCEL_AGENT_QUEUE,
  AgentJobPayload,
  AgentJobResult,
} from './job.types';

@Processor(EXCEL_AGENT_QUEUE)
export class JobsProcessor extends WorkerHost {
  private readonly logger = new Logger(JobsProcessor.name);

  constructor(
    private readonly agentService: AgentService,
    private readonly excelService: ExcelService,
    private readonly operationsService: OperationsService,
    private readonly executionRouter: OperationExecutionRouter,
    private readonly filesService: FilesService,
  ) {
    super();
  }

  async process(job: Job<AgentJobPayload>): Promise<AgentJobResult> {
    const { jobId, storedFileName, originalName, instruction, clarificationAnswer } = job.data;
    this.logger.log(`Processing job ${jobId}: "${instruction.slice(0, 80)}"`);

    try {
      // 1. Read workbook from disk (never from queue payload)
      let buffer: Buffer;
      try {
        buffer = await this.filesService.readFileFromDisk(storedFileName);
      } catch {
        return {
          status: 'failed',
          jobId,
          error: `Uploaded file not found on disk (${storedFileName}). It may have been cleaned up.`,
          finishedAt: new Date().toISOString(),
        };
      }

      // 2. Build effective instruction — include clarification answer when present
      const effectiveInstruction = clarificationAnswer
        ? `${instruction}\n\nUser clarification: ${clarificationAnswer}`
        : instruction;

      // 3. Inspect workbook
      await job.updateProgress(5);
      const workbookInfo = await this.excelService.inspectWorkbook(buffer);

      // 4. Run agent
      await job.updateProgress(10);
      const agentResult = await this.agentService.runAgent(
        buffer,
        effectiveInstruction,
        workbookInfo,
      );

      // 4. Handle non-completed statuses
      if (agentResult.status === 'clarification_needed') {
        this.logger.log(`Job ${jobId}: clarification needed`);
        return {
          status: 'clarification_needed',
          jobId,
          agentActivityLog: agentResult.activityLog,
          agentRoundsUsed: agentResult.roundsUsed,
          agentToolCallsMade: agentResult.toolCallsMade,
          clarificationQuestion: agentResult.clarificationQuestion,
          clarificationCandidates: agentResult.clarificationCandidates,
          finishedAt: new Date().toISOString(),
        };
      }

      if (agentResult.status !== 'completed' || !agentResult.plan) {
        this.logger.log(`Job ${jobId}: cannot_proceed — ${agentResult.cannotProceedReason}`);
        return {
          status: 'cannot_proceed',
          jobId,
          cannotProceedReason: agentResult.cannotProceedReason,
          agentActivityLog: agentResult.activityLog,
          agentRoundsUsed: agentResult.roundsUsed,
          agentToolCallsMade: agentResult.toolCallsMade,
          finishedAt: new Date().toISOString(),
        };
      }

      // 5. Validate + execute plan
      await job.updateProgress(70);
      const plan = agentResult.plan;
      this.operationsService.validatePlan(plan);

      const workbook = await this.excelService.loadWorkbook(buffer);
      const routerResult = await this.executionRouter.executePlan(buffer, plan, workbook);

      // 6. Save output
      await job.updateProgress(90);
      const outputBuffer = await this.excelService.saveWorkbook(workbook);
      const outputStoredName = await this.filesService.saveOutputFile(outputBuffer, originalName);

      await job.updateProgress(100);
      this.logger.log(`Job ${jobId}: completed → ${outputStoredName}`);

      return {
        status: 'completed',
        jobId,
        outputStoredName,
        downloadUrl: `/api/files/download/${outputStoredName}`,
        operationSummary: plan.summary,
        operationsApplied: plan.operations.length,
        writeOpsApplied: routerResult.writeOpsApplied,
        plan,
        agentActivityLog: agentResult.activityLog,
        agentRoundsUsed: agentResult.roundsUsed,
        agentToolCallsMade: agentResult.toolCallsMade,
        finishedAt: new Date().toISOString(),
      };

    } catch (err) {
      this.logger.error(`Job ${jobId} failed: ${err.message}`);
      return {
        status: 'failed',
        jobId,
        error: err.message ?? 'Unknown error during processing',
        finishedAt: new Date().toISOString(),
      };
    }
  }
}
