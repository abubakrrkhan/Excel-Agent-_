/**
 * operation-execution-router.ts
 *
 * Phase 2 Step 2 — Operation Execution Router.
 *
 * Sits between FilesController and ExcelService / OperationsService.
 * Routes each operation in a validated plan to the correct executor:
 *
 *   READ ops  → OperationsService inspection executors (read-only, no workbook change)
 *   HYBRID    → SET_CELL_BY_LABEL resolution (resolve only, never auto-writes)
 *   WRITE ops → ExcelService.applyOperationPlan() (modifies loaded workbook)
 *
 * Design principles:
 *   1. Read ops never modify the workbook.
 *   2. SET_CELL_BY_LABEL only resolves — never writes automatically.
 *      Ambiguous or not-found results are recorded; no write occurs.
 *   3. Write ops go through the existing ExcelService path (which includes
 *      formula overwrite protection at the execution layer).
 *   4. The original file buffer is NEVER modified.
 *   5. One router service — no duplication of inspection or Excel logic.
 */

import { Injectable, Logger } from '@nestjs/common';
import * as ExcelJS from 'exceljs';
import { ExcelOperation, OperationPlan } from '../ai/ai.service';
import { ExcelService } from '../excel/excel.service';
import { OperationsService } from '../operations/operations.service';
import { SetCellByLabelResult } from '../operations/operation.types';
import {
  SheetStructure,
  HeaderDetectionResult,
  SectionDetectionResult,
  CellContext,
} from '../excel/structure.types';

// ─── Result shapes ────────────────────────────────────────────────────────────

/** Result for a single read operation */
export interface ReadOperationResult {
  type: 'read';
  operationType: string;
  sheetName: string;
  result:
    | SheetStructure
    | HeaderDetectionResult
    | SectionDetectionResult
    | CellContext;
}

/** Result for a SET_CELL_BY_LABEL resolution (never writes automatically) */
export interface LabelResolutionResult {
  type: 'label_resolution';
  operationType: 'SET_CELL_BY_LABEL';
  label: string;
  resolution: SetCellByLabelResult;
  /**
   * writeBlocked = true when status is not 'resolved'.
   * No write is performed in any ambiguous or not-found case.
   */
  writeBlocked: boolean;
  /**
   * If resolved and a subsequent write was authorized (via allowFormulaOverwrite
   * or simple value write), the write result is recorded here.
   * null = no write was performed (ambiguous/not-found/blocked).
   */
  writeApplied: boolean;
}

/** Result for a write operation */
export interface WriteOperationResult {
  type: 'write';
  operationType: string;
  sheetName: string;
  applied: true;
}

export type SingleOperationResult =
  | ReadOperationResult
  | LabelResolutionResult
  | WriteOperationResult;

/**
 * Complete result from routing an entire operation plan.
 */
export interface ExecutionRouterResult {
  /** Number of write operations applied to the workbook */
  writeOpsApplied: number;
  /** Number of read operations executed (workbook unchanged) */
  readOpsExecuted: number;
  /** Number of SET_CELL_BY_LABEL resolutions performed */
  labelResolutionsPerformed: number;
  /** Number of SET_CELL_BY_LABEL resolutions that were blocked (ambiguous/not-found) */
  labelResolutionsBlocked: number;
  /** Per-operation results in execution order */
  operationResults: SingleOperationResult[];
  /** Whether any write operations were applied */
  workbookModified: boolean;
}

// ─── Router ───────────────────────────────────────────────────────────────────

@Injectable()
export class OperationExecutionRouter {
  private readonly logger = new Logger(OperationExecutionRouter.name);

  constructor(
    private readonly operationsService: OperationsService,
    private readonly excelService: ExcelService,
  ) {}

  /**
   * Execute a validated operation plan against a workbook buffer.
   *
   * @param buffer       The original workbook buffer (never modified)
   * @param plan         The validated operation plan
   * @param workbook     A pre-loaded ExcelJS workbook (loaded from buffer by caller)
   * @returns            Execution results and the modified workbook (if any writes occurred)
   *
   * IMPORTANT:
   *   - `buffer` is used by read ops and SET_CELL_BY_LABEL for live workbook reads.
   *   - `workbook` is mutated by write ops only.
   *   - The caller must call ExcelService.saveWorkbook(workbook) if workbookModified=true.
   */
  async executePlan(
    buffer: Buffer,
    plan: OperationPlan,
    workbook: ExcelJS.Workbook,
  ): Promise<ExecutionRouterResult> {
    const operationResults: SingleOperationResult[] = [];
    let writeOpsApplied = 0;
    let readOpsExecuted = 0;
    let labelResolutionsPerformed = 0;
    let labelResolutionsBlocked = 0;

    // Collect write ops to batch-apply through existing ExcelService path
    const pendingWriteOps: ExcelOperation[] = [];

    for (const op of plan.operations) {
      if (this.operationsService.isReadOperation(op.type)) {
        // ── READ ──────────────────────────────────────────────────────────
        const result = await this.executeReadOp(buffer, op);
        operationResults.push(result);
        readOpsExecuted++;

        this.logger.log(
          `[READ] ${op.type} on "${op.sheet}" → completed`,
        );
      } else if (op.type === 'SET_CELL_BY_LABEL') {
        // ── HYBRID: flush any pending writes first, then resolve ──────────
        if (pendingWriteOps.length > 0) {
          const batchResult = this.flushWriteOps(workbook, pendingWriteOps, plan.summary);
          operationResults.push(...batchResult);
          writeOpsApplied += batchResult.length;
          pendingWriteOps.length = 0;
        }

        const labelResult = await this.executeLabelResolution(
          buffer,
          workbook,
          op,
        );
        operationResults.push(labelResult);
        labelResolutionsPerformed++;

        if (labelResult.writeBlocked) {
          labelResolutionsBlocked++;
          this.logger.warn(
            `[SET_CELL_BY_LABEL] label="${op.params?.label}" → ` +
              `status=${labelResult.resolution.status} — write BLOCKED`,
          );
        } else {
          this.logger.log(
            `[SET_CELL_BY_LABEL] label="${op.params?.label}" → ` +
              `resolved to ${labelResult.resolution.resolvedTarget?.valueCoordinate} on ` +
              `"${labelResult.resolution.resolvedTarget?.sheetName}"` +
              (labelResult.writeApplied ? ', value written' : ', write deferred'),
          );
        }
      } else {
        // ── WRITE ─────────────────────────────────────────────────────────
        pendingWriteOps.push(op);
      }
    }

    // Flush any remaining write ops
    if (pendingWriteOps.length > 0) {
      const batchResult = this.flushWriteOps(workbook, pendingWriteOps, plan.summary);
      operationResults.push(...batchResult);
      writeOpsApplied += batchResult.length;
    }

    const workbookModified = writeOpsApplied > 0;

    this.logger.log(
      `Router complete: ${writeOpsApplied} write(s), ${readOpsExecuted} read(s), ` +
        `${labelResolutionsPerformed} label resolution(s) ` +
        `(${labelResolutionsBlocked} blocked), modified=${workbookModified}`,
    );

    return {
      writeOpsApplied,
      readOpsExecuted,
      labelResolutionsPerformed,
      labelResolutionsBlocked,
      operationResults,
      workbookModified,
    };
  }

  // ─── Private: read ops ────────────────────────────────────────────────────

  private async executeReadOp(
    buffer: Buffer,
    op: ExcelOperation,
  ): Promise<ReadOperationResult> {
    let result: ReadOperationResult['result'];

    switch (op.type) {
      case 'INSPECT_WORKBOOK_STRUCTURE':
        result = await this.operationsService.executeInspectStructure(
          buffer,
          op.params,
        );
        break;
      case 'DETECT_HEADERS':
        result = await this.operationsService.executeDetectHeaders(
          buffer,
          op.params,
        );
        break;
      case 'DETECT_SECTIONS':
        result = await this.operationsService.executeDetectSections(
          buffer,
          op.params,
        );
        break;
      case 'GET_CELL_CONTEXT':
        result = await this.operationsService.executeGetCellContext(
          buffer,
          op.params,
        );
        break;
      default:
        throw new Error(
          `OperationExecutionRouter: unhandled read op "${op.type}"`,
        );
    }

    return {
      type: 'read',
      operationType: op.type,
      sheetName: op.sheet,
      result,
    };
  }

  // ─── Private: SET_CELL_BY_LABEL ───────────────────────────────────────────

  /**
   * Resolve a SET_CELL_BY_LABEL operation.
   *
   * Resolution rules (anti-hallucination):
   *   - status = 'not_found'   → writeBlocked=true, no write
   *   - status = 'ambiguous'   → writeBlocked=true, no write (NEVER auto-selects)
   *   - status = 'resolved'    → apply the value to resolvedTarget.valueCoordinate
   *
   * The write is executed directly on the already-loaded workbook so we don't
   * need to reload it. Formula overwrite protection still applies through
   * ExcelService.applyOperationPlan().
   */
  private async executeLabelResolution(
    buffer: Buffer,
    workbook: ExcelJS.Workbook,
    op: ExcelOperation,
  ): Promise<LabelResolutionResult> {
    const resolution = await this.operationsService.executeSetCellByLabel(
      buffer,
      op.params,
    );

    // Only resolved status allows a write
    if (resolution.status !== 'resolved' || !resolution.resolvedTarget) {
      return {
        type: 'label_resolution',
        operationType: 'SET_CELL_BY_LABEL',
        label: String(op.params?.label ?? ''),
        resolution,
        writeBlocked: true,
        writeApplied: false,
      };
    }

    // Write the value to the resolved coordinate
    const target = resolution.resolvedTarget;
    const writeOp: ExcelOperation = {
      type: 'SET_CELL',
      sheet: target.sheetName,
      params: {
        // Parse row/col from the A1 coordinate
        ...this.coordToRowCol(target.valueCoordinate),
        value: op.params?.value ?? null,
        allowFormulaOverwrite: op.params?.allowFormulaOverwrite ?? false,
      },
    };

    // Use ExcelService to apply the write (formula protection active)
    const writePlan = { summary: `SET_CELL_BY_LABEL write: ${target.valueCoordinate}`, operations: [writeOp] };
    this.excelService.applyOperationPlan(workbook, writePlan);

    return {
      type: 'label_resolution',
      operationType: 'SET_CELL_BY_LABEL',
      label: String(op.params?.label ?? ''),
      resolution,
      writeBlocked: false,
      writeApplied: true,
    };
  }

  // ─── Private: write ops ───────────────────────────────────────────────────

  /**
   * Flush a batch of write ops through ExcelService.applyOperationPlan().
   * This preserves the existing formula protection and validation.
   */
  private flushWriteOps(
    workbook: ExcelJS.Workbook,
    ops: ExcelOperation[],
    summary: string,
  ): WriteOperationResult[] {
    const batchPlan = { summary, operations: ops };
    this.excelService.applyOperationPlan(workbook, batchPlan);

    return ops.map((op) => ({
      type: 'write' as const,
      operationType: op.type,
      sheetName: op.sheet,
      applied: true as const,
    }));
  }

  // ─── Private: coordinate parsing ─────────────────────────────────────────

  /**
   * Convert an A1 coordinate to { row, col } for SET_CELL params.
   */
  private coordToRowCol(coordinate: string): { row: number; col: number } {
    const { parseCell } = require('../excel/utils/coordinate.utils');
    const parsed = parseCell(coordinate);
    return { row: parsed.row, col: parsed.col };
  }
}
