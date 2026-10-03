import {
  Controller,
  Post,
  Get,
  Param,
  UploadedFile,
  UploadedFiles,
  UseInterceptors,
  HttpCode,
  HttpStatus,
  Logger,
  Res,
  Body,
} from '@nestjs/common';
import { FileInterceptor, FileFieldsInterceptor } from '@nestjs/platform-express';
import { memoryStorage } from 'multer';
import * as crypto from 'crypto';
import * as path from 'path';
import type { Response } from 'express';
import { FilesService } from './files.service';
import { ExcelService, WorkbookInfo } from '../excel/excel.service';
import { AiService, OperationPlan } from '../ai/ai.service';
import { OperationsService } from '../operations/operations.service';
import { OperationExecutionRouter } from '../operations/operation-execution-router';
import { AgentService } from '../agent/agent.service';
import { AgentExecutionResult } from '../agent/agent.types';

// ─── Response shapes ──────────────────────────────────────────────────────────

interface UploadResponse {
  id: string;
  originalName: string;
  mimeType: string;
  sizeBytes: number;
  status: string;
  uploadedAt: string | null;
  workbook: WorkbookInfo;
}

interface FileMetaResponse {
  id: string;
  originalName: string;
  mimeType: string;
  sizeBytes: number;
  status: string;
  uploadedAt: string | null;
  updatedAt: string | null;
}

interface ProcessResponse {
  status: 'success' | 'clarification_needed' | 'cannot_proceed' | 'max_rounds_reached';
  originalName: string;
  // Present only when status = 'success'
  outputFileId?: string;
  outputStoredName?: string;
  downloadUrl?: string;
  workbookSummary: WorkbookInfo;
  operationSummary: string;
  operationsApplied: number;
  writeOpsApplied: number;
  readOpsExecuted: number;
  labelResolutionsPerformed: number;
  labelResolutionsBlocked: number;
  plan: OperationPlan | null;
  // Agent fields
  agentRoundsUsed: number;
  agentToolCallsMade: number;
  agentActivityLog: string[];
  // Present only when status = 'clarification_needed'
  clarificationQuestion?: string;
  clarificationCandidates?: string[];
  // Present only when status = 'cannot_proceed' or 'max_rounds_reached'
  cannotProceedReason?: string;
}

// ─── Controller ───────────────────────────────────────────────────────────────

const MULTER_OPTIONS = {
  storage: memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024 }, // 10MB
};

@Controller('files')
export class FilesController {
  private readonly logger = new Logger(FilesController.name);

  constructor(
    private readonly filesService: FilesService,
    private readonly excelService: ExcelService,
    private readonly aiService: AiService,
    private readonly operationsService: OperationsService,
    private readonly executionRouter: OperationExecutionRouter,
    private readonly agentService: AgentService,
  ) {}

  // ── POST /api/files/upload ─────────────────────────────────────────────────
  //
  // NOTE: Clarification is handled by POST /api/jobs/submit with the optional
  // clarificationAnswer body field (async path through BullMQ). The synchronous
  // POST /api/files/clarify endpoint was removed in Phase 11 because it had no
  // frontend references and ran the agent inline, risking HTTP timeout.

  @Post('upload')
  @HttpCode(HttpStatus.CREATED)
  @UseInterceptors(FileInterceptor('file', MULTER_OPTIONS))
  async uploadFile(
    @UploadedFile() file: Express.Multer.File,
  ): Promise<UploadResponse> {
    this.filesService.validateExcelFile(file);
    const workbookInfo = await this.excelService.inspectWorkbook(file.buffer);
    const storedName = await this.filesService.saveFileToDisk(file);

    const id = crypto.randomUUID();
    const record = await this.filesService.createFileRecord({
      id,
      originalName: file.originalname,
      storedName,
      mimeType: file.mimetype,
      sizeBytes: file.size,
    });

    this.logger.log(
      `Upload complete: "${file.originalname}" → id=${id}, sheets=${workbookInfo.sheetCount}`,
    );

    return {
      id: record?.id ?? id,
      originalName: file.originalname,
      mimeType: file.mimetype,
      sizeBytes: file.size,
      status: record?.status ?? 'PENDING',
      uploadedAt: record?.uploadedAt?.toISOString() ?? null,
      workbook: workbookInfo,
    };
  }

  // ── POST /api/files/process ────────────────────────────────────────────────

  /**
   * Complete end-to-end workflow:
   *   1. Accept file + instruction
   *   2. Validate file
   *   3. Inspect workbook → send to AI
   *   4. AI generates OperationPlan
   *   5. OperationsService validates plan
   *   6. ExcelService applies plan
   *   7. Save output file (never overwrites original)
   *   8. Return result + download URL
   */
  @Post('process')
  @HttpCode(HttpStatus.OK)
  @UseInterceptors(
    FileFieldsInterceptor(
      [{ name: 'file', maxCount: 1 }],
      MULTER_OPTIONS,
    ),
  )
  async processFile(
    @UploadedFiles() files: { file?: Express.Multer.File[] },
    @Body('instruction') instruction: string,
  ): Promise<ProcessResponse> {
    const file = files?.file?.[0];

    // ── Step 1: validate file ──────────────────────────────────────────────
    this.filesService.validateExcelFile(file);

    // ── Step 2: validate instruction ──────────────────────────────────────
    // AiService.validateInputs will also check, but we want a clean early error
    if (!instruction || !instruction.trim()) {
      const { BadRequestException } = await import('@nestjs/common');
      throw new BadRequestException('Instruction is required and cannot be empty');
    }

    this.logger.log(
      `Process request: "${file.originalname}" | instruction="${instruction.slice(0, 80)}"`,
    );

    // ── Step 3: inspect workbook ───────────────────────────────────────────
    const workbookInfo = await this.excelService.inspectWorkbook(file.buffer);

    // ── Step 4: run the AI agent (multi-round tool-calling loop) ──────────
    //    Agent inspects the workbook, gathers evidence, then produces a plan.
    //    Falls back to generateOperationPlan for non-agent paths if needed.
    const agentResult = await this.agentService.runAgent(
      file.buffer,
      instruction,
      workbookInfo,
    );

    // ── Step 5: handle non-success agent outcomes ──────────────────────────
    if (agentResult.status === 'clarification_needed') {
      this.logger.log(
        `Agent needs clarification: "${agentResult.clarificationQuestion}"`,
      );
      return {
        status: 'clarification_needed',
        originalName: file.originalname,
        workbookSummary: workbookInfo,
        operationSummary: 'Clarification required before proceeding',
        operationsApplied: 0,
        writeOpsApplied: 0,
        readOpsExecuted: agentResult.toolCallsMade,
        labelResolutionsPerformed: 0,
        labelResolutionsBlocked: 0,
        plan: null,
        agentRoundsUsed: agentResult.roundsUsed,
        agentToolCallsMade: agentResult.toolCallsMade,
        agentActivityLog: agentResult.activityLog,
        clarificationQuestion: agentResult.clarificationQuestion,
        clarificationCandidates: agentResult.clarificationCandidates,
      };
    }

    if (agentResult.status === 'cannot_proceed' || agentResult.status === 'max_rounds_reached') {
      this.logger.log(
        `Agent could not complete: status=${agentResult.status}`,
      );
      return {
        status: agentResult.status,
        originalName: file.originalname,
        workbookSummary: workbookInfo,
        operationSummary: agentResult.cannotProceedReason ?? 'Agent could not produce a safe plan',
        operationsApplied: 0,
        writeOpsApplied: 0,
        readOpsExecuted: agentResult.toolCallsMade,
        labelResolutionsPerformed: 0,
        labelResolutionsBlocked: 0,
        plan: null,
        agentRoundsUsed: agentResult.roundsUsed,
        agentToolCallsMade: agentResult.toolCallsMade,
        agentActivityLog: agentResult.activityLog,
        cannotProceedReason: agentResult.cannotProceedReason,
      };
    }

    // ── Step 6: agent produced a plan — validate (defence layer) ──────────
    const plan = agentResult.plan!;
    this.operationsService.validatePlan(plan);

    // ── Step 7: route operations through the execution router ─────────────
    const workbook = await this.excelService.loadWorkbook(file.buffer);
    const routerResult = await this.executionRouter.executePlan(
      file.buffer,
      plan,
      workbook,
    );

    // ── Step 8: save output ────────────────────────────────────────────────
    const outputBuffer = await this.excelService.saveWorkbook(workbook);
    const outputStoredName = await this.filesService.saveOutputFile(
      outputBuffer,
      file.originalname,
    );
    const outputFileId = crypto.randomUUID();

    this.logger.log(
      `Process complete: ${routerResult.writeOpsApplied} write(s), ` +
        `${agentResult.roundsUsed} agent rounds, ` +
        `${agentResult.toolCallsMade} tool calls → ${outputStoredName}`,
    );

    return {
      status: 'success',
      originalName: file.originalname,
      outputFileId,
      outputStoredName,
      downloadUrl: `/api/files/download/${outputStoredName}`,
      workbookSummary: workbookInfo,
      operationSummary: plan.summary,
      operationsApplied: plan.operations.length,
      writeOpsApplied: routerResult.writeOpsApplied,
      readOpsExecuted: routerResult.readOpsExecuted,
      labelResolutionsPerformed: routerResult.labelResolutionsPerformed,
      labelResolutionsBlocked: routerResult.labelResolutionsBlocked,
      plan,
      agentRoundsUsed: agentResult.roundsUsed,
      agentToolCallsMade: agentResult.toolCallsMade,
      agentActivityLog: agentResult.activityLog,
    };
  }

  // ── GET /api/files/:id ─────────────────────────────────────────────────────

  @Get(':id')
  async getFile(@Param('id') id: string): Promise<FileMetaResponse> {
    const record = await this.filesService.findFileRecord(id);

    return {
      id: record.id,
      originalName: record.originalName,
      mimeType: record.mimeType,
      sizeBytes: record.sizeBytes,
      status: record.status,
      uploadedAt: record.uploadedAt
        ? new Date(record.uploadedAt).toISOString()
        : null,
      updatedAt: record.updatedAt
        ? new Date(record.updatedAt).toISOString()
        : null,
    };
  }

  // ── GET /api/files/download/:filename ─────────────────────────────────────

  /**
   * Download a processed output file by its stored filename.
   *
   * Uses stored filename (UUID-based) not an arbitrary path.
   * path.basename() prevents directory traversal.
   * Returns the file as application/xlsx with a sensible download name.
   */
  @Get('download/:filename')
  async downloadFile(
    @Param('filename') filename: string,
    @Res() res: Response,
  ): Promise<void> {
    // Security: strip any directory components (path traversal prevention)
    const safeFilename = path.basename(filename);

    // Security: only allow output files (UUID-named outputs prefixed with 'output-')
    // This prevents downloading uploaded input files or other arbitrary files.
    if (!safeFilename.startsWith('output-')) {
      const { ForbiddenException } = await import('@nestjs/common');
      throw new ForbiddenException('Only processed output files can be downloaded.');
    }

    // Security: only allow valid UUID-based filenames (no special characters)
    if (!/^output-[0-9a-f-]+\.(xlsx)$/.test(safeFilename)) {
      const { BadRequestException } = await import('@nestjs/common');
      throw new BadRequestException('Invalid filename format.');
    }

    if (!this.filesService.fileExistsOnDisk(safeFilename)) {
      const { NotFoundException } = await import('@nestjs/common');
      throw new NotFoundException(
        `Output file not found or has expired: ${safeFilename}`,
      );
    }

    const filePath = this.filesService.getOutputFilePath(safeFilename);

    // Strip the output- prefix and uuid for a cleaner download name
    const downloadName = safeFilename.startsWith('output-')
      ? `processed-${safeFilename.replace(/^output-[0-9a-f-]+/, '').replace(/^-/, '') || 'workbook.xlsx'}`
      : safeFilename;

    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );
    res.setHeader(
      'Content-Disposition',
      `attachment; filename="${downloadName || 'processed.xlsx'}"`,
    );

    res.sendFile(filePath, (err) => {
      if (err) {
        this.logger.error(`sendFile failed for ${safeFilename}`, err.message);
        if (!res.headersSent) {
          res.status(500).json({ message: 'Failed to stream output file' });
        }
      }
    });
  }
}
