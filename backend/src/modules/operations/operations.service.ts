import { Injectable, Logger, BadRequestException } from '@nestjs/common';
import { ExcelOperation, OperationPlan } from '../ai/ai.service';
import { ExcelInspectionService } from '../excel/excel-inspection.service';
import { WorkbookStructureService } from '../excel/workbook-structure.service';
import { isCellAddressValid } from '../excel/utils/coordinate.utils';
import {
  SheetStructure,
  HeaderDetectionResult,
  SectionDetectionResult,
  CellContext,
  CellContextOptions,
} from '../excel/structure.types';
import {
  ALL_OPERATION_TYPES,
  READ_OPERATIONS,
  SetCellByLabelParams,
  SetCellByLabelResult,
  LabelMatch,
} from './operation.types';

// ─── Operation type sets ──────────────────────────────────────────────────────

const WRITE_OPS = new Set([
  'SET_CELL', 'SET_CELL_RANGE', 'ADD_ROW', 'DELETE_ROW',
  'ADD_FORMULA', 'FORMAT_CELL', 'ADD_COLUMN_HEADER', 'SORT_SHEET',
]);

const READ_OPS = new Set(READ_OPERATIONS as readonly string[]);

// ─── Service ──────────────────────────────────────────────────────────────────

@Injectable()
export class OperationsService {
  private readonly logger = new Logger(OperationsService.name);

  constructor(
    private readonly inspectionService: ExcelInspectionService,
    private readonly structureService: WorkbookStructureService,
  ) {}

  // ─── Plan validation ──────────────────────────────────────────────────────

  /**
   * Validate an operation plan before execution.
   *
   * Validates:
   *   - plan structure
   *   - each operation type is in ALL_OPERATION_TYPES
   *   - required params present and correctly typed
   *   - formula safety params on SET_CELL
   *
   * Does NOT execute operations.
   * Does NOT load the workbook.
   * For write ops, formula-overwrite protection is enforced at both validation
   * (structural check) and execution (live cell check) layers.
   */
  validatePlan(plan: OperationPlan): void {
    if (!plan || !Array.isArray(plan.operations)) {
      throw new BadRequestException('Invalid operation plan');
    }
    if (plan.operations.length === 0) {
      throw new BadRequestException('Operation plan is empty');
    }
    for (const op of plan.operations) {
      this.validateOperation(op);
    }
    this.logger.log(`Validated ${plan.operations.length} operation(s)`);
  }

  /**
   * Returns true if the operation type is a read-only inspection operation.
   */
  isReadOperation(type: string): boolean {
    return READ_OPS.has(type);
  }

  /**
   * Returns true if the operation type modifies the workbook.
   */
  isWriteOperation(type: string): boolean {
    return WRITE_OPS.has(type);
  }

  // ─── Inspection operations ────────────────────────────────────────────────

  /**
   * Execute INSPECT_WORKBOOK_STRUCTURE.
   * Returns the existing SheetStructure from WorkbookStructureService.
   */
  async executeInspectStructure(
    buffer: Buffer,
    params: Record<string, any>,
  ): Promise<SheetStructure> {
    this.validateInspectionSheetParam(params);
    return this.structureService.inspectSheetStructure(buffer, params.sheetName as string);
  }

  /**
   * Execute DETECT_HEADERS.
   * Returns the existing HeaderDetectionResult from WorkbookStructureService.
   */
  async executeDetectHeaders(
    buffer: Buffer,
    params: Record<string, any>,
  ): Promise<HeaderDetectionResult> {
    this.validateInspectionSheetParam(params);
    return this.structureService.detectHeaders(buffer, params.sheetName as string);
  }

  /**
   * Execute DETECT_SECTIONS.
   * Returns the existing SectionDetectionResult from WorkbookStructureService.
   */
  async executeDetectSections(
    buffer: Buffer,
    params: Record<string, any>,
  ): Promise<SectionDetectionResult> {
    this.validateInspectionSheetParam(params);
    return this.structureService.detectSections(buffer, params.sheetName as string);
  }

  /**
   * Execute GET_CELL_CONTEXT.
   * Returns the existing CellContext from WorkbookStructureService.
   */
  async executeGetCellContext(
    buffer: Buffer,
    params: Record<string, any>,
  ): Promise<CellContext> {
    this.validateCellContextParams(params);
    const options: CellContextOptions = {};
    if (typeof params.rowRadius === 'number') options.rowRadius = params.rowRadius;
    if (typeof params.colRadius === 'number') options.colRadius = params.colRadius;
    return this.structureService.getCellContext(
      buffer,
      params.sheetName as string,
      params.coordinate as string,
      options,
    );
  }

  // ─── SET_CELL_BY_LABEL ────────────────────────────────────────────────────

  /**
   * Execute SET_CELL_BY_LABEL — evidence-based label-to-cell resolution.
   *
   * Flow:
   *   1. Search the workbook for the label text
   *   2. Apply offset to find the value cell
   *   3. If exactly one match → return 'resolved' with target coordinate
   *   4. If multiple matches → return 'ambiguous' (never auto-select)
   *   5. If no matches → return 'not_found'
   *
   * IMPORTANT: This method ONLY resolves the target. It does NOT write.
   * The caller must use the returned resolvedTarget.valueCoordinate
   * in a subsequent SET_CELL operation after the user/agent confirms.
   *
   * This enforces the Evidence Rule: a cell cannot be written without
   * direct workbook inspection evidence from the current execution.
   */
  async executeSetCellByLabel(
    buffer: Buffer,
    params: Record<string, any>,
  ): Promise<SetCellByLabelResult> {
    this.validateSetCellByLabelParams(params);

    const label = params.label as string;
    const valueColOffset = typeof params.valueColOffset === 'number'
      ? params.valueColOffset
      : 1;
    const sheetFilter = typeof params.sheetName === 'string' && params.sheetName.trim()
      ? params.sheetName.trim()
      : undefined;

    // 1. Search the workbook for all matches
    const searchResult = await this.inspectionService.searchWorkbook(
      buffer,
      label,
      {
        caseInsensitive: true,
        sheetName: sheetFilter,
        maxMatches: 20, // enough to detect ambiguity without being unbounded
      },
    );

    // 2. No matches
    if (searchResult.matches.length === 0) {
      this.logger.log(`SET_CELL_BY_LABEL: label "${label}" not found`);
      return {
        status: 'not_found',
        message: `Label "${label}" was not found in the workbook${sheetFilter ? ` on sheet "${sheetFilter}"` : ''}.`,
        candidates: [],
        resolvedTarget: null,
      };
    }

    // 3. Build candidate list with value cell coordinates
    const candidates: LabelMatch[] = [];
    for (const match of searchResult.matches) {
      const labelCell = match.coordinate;
      const valueCoord = this.applyColOffset(labelCell, valueColOffset);
      if (valueCoord) {
        candidates.push({
          sheetName: match.sheetName,
          labelCoordinate: labelCell,
          valueCoordinate: valueCoord,
          labelValue: match.value,
        });
      }
    }

    // 4. Multiple matches → ambiguous, NEVER auto-select
    if (candidates.length > 1) {
      const locations = candidates
        .map((c) => `"${c.sheetName}"!${c.labelCoordinate}`)
        .join(', ');
      this.logger.warn(
        `SET_CELL_BY_LABEL: ambiguous — "${label}" found in ${candidates.length} locations: ${locations}`,
      );
      return {
        status: 'ambiguous',
        message:
          `Ambiguous: label "${label}" was found in ${candidates.length} locations: ${locations}. ` +
          `Cannot safely resolve to a single target. ` +
          `Use GET_CELL_CONTEXT to inspect each candidate and identify the correct one.`,
        candidates,
        resolvedTarget: null,
      };
    }

    // 5. Exactly one match → resolved
    const resolved = candidates[0];
    this.logger.log(
      `SET_CELL_BY_LABEL: resolved "${label}" → ` +
        `"${resolved.sheetName}"!${resolved.valueCoordinate}`,
    );
    return {
      status: 'resolved',
      message:
        `Label "${label}" resolved to ${resolved.valueCoordinate} on sheet "${resolved.sheetName}". ` +
        `Value cell is ${resolved.valueCoordinate} (offset ${valueColOffset} from label at ${resolved.labelCoordinate}).`,
      candidates,
      resolvedTarget: resolved,
    };
  }

  // ─── Private: validation ──────────────────────────────────────────────────

  private validateOperation(op: ExcelOperation): void {
    if (!op.type) {
      throw new BadRequestException('Operation missing type');
    }
    if (!ALL_OPERATION_TYPES.includes(op.type)) {
      throw new BadRequestException(
        `Unsupported operation type: "${op.type}". ` +
          `Supported: ${ALL_OPERATION_TYPES.join(', ')}`,
      );
    }
    if (!op.sheet || typeof op.sheet !== 'string') {
      throw new BadRequestException(`Operation ${op.type} missing sheet name`);
    }
    if (!op.params || typeof op.params !== 'object') {
      throw new BadRequestException(`Operation ${op.type} missing params`);
    }

    switch (op.type) {
      // ── Write ops ──────────────────────────────────────────────────────
      case 'SET_CELL':
        this.validateSetCellParams(op.params);
        break;
      case 'ADD_ROW':
        if (!Array.isArray(op.params.values)) {
          throw new BadRequestException('ADD_ROW requires values array');
        }
        break;
      case 'DELETE_ROW':
        if (typeof op.params.row !== 'number' || !Number.isInteger(op.params.row) || op.params.row < 1) {
          throw new BadRequestException('DELETE_ROW requires a positive integer row number');
        }
        break;
      case 'ADD_FORMULA':
        if (!op.params.row || !op.params.col || !op.params.formula) {
          throw new BadRequestException('ADD_FORMULA requires row, col, and formula');
        }
        if (typeof op.params.formula !== 'string' || !op.params.formula.trim()) {
          throw new BadRequestException('ADD_FORMULA: formula must be a non-empty string');
        }
        break;
      case 'FORMAT_CELL':
        if (!op.params.row || !op.params.col) {
          throw new BadRequestException('FORMAT_CELL requires row and col');
        }
        if (!op.params.style || typeof op.params.style !== 'object') {
          throw new BadRequestException('FORMAT_CELL requires a style object');
        }
        break;
      case 'ADD_COLUMN_HEADER':
        if (typeof op.params.col !== 'number' || !Number.isInteger(op.params.col) || op.params.col < 1) {
          throw new BadRequestException('ADD_COLUMN_HEADER requires a positive integer col number');
        }
        if (!op.params.header || typeof op.params.header !== 'string') {
          throw new BadRequestException('ADD_COLUMN_HEADER requires a non-empty header string');
        }
        break;
      case 'SORT_SHEET':
        if (typeof op.params.col !== 'number' || !Number.isInteger(op.params.col) || op.params.col < 1) {
          throw new BadRequestException('SORT_SHEET requires a positive integer col number');
        }
        if (op.params.order !== 'asc' && op.params.order !== 'desc') {
          throw new BadRequestException('SORT_SHEET order must be "asc" or "desc"');
        }
        break;
      case 'SET_CELL_RANGE':
        this.validateSetCellRangeParams(op.params);
        break;
      // ── Inspection ops ─────────────────────────────────────────────────
      case 'INSPECT_WORKBOOK_STRUCTURE':
      case 'DETECT_HEADERS':
      case 'DETECT_SECTIONS':
        this.validateInspectionSheetParam(op.params);
        break;
      case 'GET_CELL_CONTEXT':
        this.validateCellContextParams(op.params);
        break;
      // ── Hybrid ops ─────────────────────────────────────────────────────
      case 'SET_CELL_BY_LABEL':
        this.validateSetCellByLabelParams(op.params);
        break;
    }
  }

  /**
   * Validate SET_CELL params including formula overwrite flag.
   *
   * Formula overwrite protection at validation level:
   *   - 'allowFormulaOverwrite' defaults to false when absent
   *   - When false (default), we cannot pre-check for formulas here
   *     because we don't have the workbook buffer at validation time.
   *     The execution layer performs the live formula check.
   *   - This validation ensures the param is correctly typed when present.
   */
  private validateSetCellParams(params: Record<string, any>): void {
    if (!params.row || !params.col) {
      throw new BadRequestException('SET_CELL requires row and col');
    }
    if (
      params.allowFormulaOverwrite !== undefined &&
      typeof params.allowFormulaOverwrite !== 'boolean'
    ) {
      throw new BadRequestException(
        'SET_CELL: allowFormulaOverwrite must be a boolean when provided',
      );
    }
  }

  /**
   * Validate SET_CELL_RANGE params including formula overwrite flag.
   * Mirrors SET_CELL validation — consistent policy across both write operations.
   */
  private validateSetCellRangeParams(params: Record<string, any>): void {
    if (!params.startRow || !params.startCol || !params.endRow || !params.endCol) {
      throw new BadRequestException(
        'SET_CELL_RANGE requires startRow, startCol, endRow, and endCol',
      );
    }
    if (
      params.allowFormulaOverwrite !== undefined &&
      typeof params.allowFormulaOverwrite !== 'boolean'
    ) {
      throw new BadRequestException(
        'SET_CELL_RANGE: allowFormulaOverwrite must be a boolean when provided',
      );
    }
  }

  private validateInspectionSheetParam(params: Record<string, any>): void {
    if (!params.sheetName || typeof params.sheetName !== 'string' || !params.sheetName.trim()) {
      throw new BadRequestException(
        'Inspection operation requires a non-empty sheetName parameter',
      );
    }
  }

  private validateCellContextParams(params: Record<string, any>): void {
    this.validateInspectionSheetParam(params);
    if (!params.coordinate || typeof params.coordinate !== 'string') {
      throw new BadRequestException(
        'GET_CELL_CONTEXT requires a coordinate parameter (e.g. "G42")',
      );
    }
    if (!isCellAddressValid(params.coordinate)) {
      throw new BadRequestException(
        `GET_CELL_CONTEXT: invalid coordinate "${params.coordinate}"`,
      );
    }
    if (params.rowRadius !== undefined) {
      const rr = Number(params.rowRadius);
      if (!Number.isInteger(rr) || rr < 0 || rr > 20) {
        throw new BadRequestException(
          'GET_CELL_CONTEXT: rowRadius must be an integer between 0 and 20',
        );
      }
    }
    if (params.colRadius !== undefined) {
      const cr = Number(params.colRadius);
      if (!Number.isInteger(cr) || cr < 0 || cr > 20) {
        throw new BadRequestException(
          'GET_CELL_CONTEXT: colRadius must be an integer between 0 and 20',
        );
      }
    }
  }

  private validateSetCellByLabelParams(params: Record<string, any>): void {
    if (!params.label || typeof params.label !== 'string' || !params.label.trim()) {
      throw new BadRequestException(
        'SET_CELL_BY_LABEL requires a non-empty label parameter',
      );
    }
    if (params.label.trim().length > 500) {
      throw new BadRequestException(
        'SET_CELL_BY_LABEL: label must not exceed 500 characters',
      );
    }
    if (params.valueColOffset !== undefined) {
      const offset = Number(params.valueColOffset);
      if (!Number.isInteger(offset) || offset < -50 || offset > 50) {
        throw new BadRequestException(
          'SET_CELL_BY_LABEL: valueColOffset must be an integer between -50 and 50',
        );
      }
    }
    if (
      params.allowFormulaOverwrite !== undefined &&
      typeof params.allowFormulaOverwrite !== 'boolean'
    ) {
      throw new BadRequestException(
        'SET_CELL_BY_LABEL: allowFormulaOverwrite must be a boolean when provided',
      );
    }
  }

  /**
   * Apply a column offset to an A1 address.
   * Returns null if the result would be an invalid coordinate.
   */
  private applyColOffset(coordinate: string, offset: number): string | null {
    try {
      const { parseCell, buildCellAddress } = require('../excel/utils/coordinate.utils');
      const parsed = parseCell(coordinate);
      const newCol = parsed.col + offset;
      if (newCol < 1) return null;
      return buildCellAddress(parsed.row, newCol);
    } catch {
      return null;
    }
  }
}
