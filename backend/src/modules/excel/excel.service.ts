import {
  Injectable,
  Logger,
  BadRequestException,
  InternalServerErrorException,
} from '@nestjs/common';
import * as ExcelJS from 'exceljs';
import { OperationPlan, ExcelOperation } from '../ai/ai.service';

// ─── Types ────────────────────────────────────────────────────────────────────

export interface CellValue {
  raw: string | number | boolean | null;
  type: 'string' | 'number' | 'boolean' | 'date' | 'formula' | 'empty';
}

export interface SheetInfo {
  name: string;
  /** Total rows including empty trailing rows reported by ExcelJS */
  rowCount: number;
  /** Rows that actually contain data */
  dataRowCount: number;
  /** Total columns reported by ExcelJS */
  columnCount: number;
  /** Columns that actually contain data */
  dataColumnCount: number;
  /**
   * Header values from row 1, one entry per column.
   * Empty cells are represented as empty strings so column positions are preserved.
   */
  headers: string[];
}

export interface WorkbookInfo {
  sheetCount: number;
  sheets: SheetInfo[];
}

// ─── Service ──────────────────────────────────────────────────────────────────

@Injectable()
export class ExcelService {
  private readonly logger = new Logger(ExcelService.name);

  // ── Inspection ──────────────────────────────────────────────────────────────

  /**
   * Load an .xlsx Buffer and return structured workbook metadata.
   * The AI receives only this metadata — never the file itself.
   */
  async inspectWorkbook(buffer: Buffer): Promise<WorkbookInfo> {
    const workbook = await this.loadWorkbook(buffer);
    const sheets: SheetInfo[] = [];

    workbook.eachSheet((worksheet) => {
      sheets.push(this.inspectSheet(worksheet));
    });

    this.logger.log(
      `Inspected workbook: ${sheets.length} sheet(s) — ` +
        sheets.map((s) => `"${s.name}" (${s.dataRowCount} rows)`).join(', '),
    );

    return { sheetCount: sheets.length, sheets };
  }

  // ── Load / Save ─────────────────────────────────────────────────────────────

  /**
   * Load a workbook from a Buffer.
   */
  async loadWorkbook(buffer: Buffer): Promise<ExcelJS.Workbook> {
    if (!buffer || buffer.length === 0) {
      throw new BadRequestException('File buffer is empty');
    }

    try {
      const workbook = new ExcelJS.Workbook();
      // ExcelJS declares its own Buffer interface that conflicts with Node's Buffer.
      // Cast is required at the call site only.
      await workbook.xlsx.load(
        buffer as unknown as Parameters<typeof workbook.xlsx.load>[0],
      );
      return workbook;
    } catch (err) {
      this.logger.error('Failed to load workbook', err.message);
      throw new BadRequestException(
        'Could not read Excel file — make sure it is a valid .xlsx file',
      );
    }
  }

  /**
   * Serialize a workbook to a Buffer for download or disk write.
   */
  async saveWorkbook(workbook: ExcelJS.Workbook): Promise<Buffer> {
    const arrayBuffer = await workbook.xlsx.writeBuffer();
    return Buffer.from(arrayBuffer);
  }

  // ── Operation Execution ─────────────────────────────────────────────────────

  /**
   * Apply a validated OperationPlan to an ExcelJS Workbook.
   *
   * IMPORTANT: This method expects an ALREADY VALIDATED plan.
   * Call OperationsService.validatePlan() before passing the plan here.
   * Operations are applied in order. Errors in individual operations throw
   * immediately and stop processing.
   *
   * The workbook is modified in-place. Call saveWorkbook() afterwards.
   */
  applyOperationPlan(
    workbook: ExcelJS.Workbook,
    plan: OperationPlan,
  ): void {
    this.logger.log(
      `Applying ${plan.operations.length} operation(s): ${plan.summary}`,
    );

    for (let i = 0; i < plan.operations.length; i++) {
      const op = plan.operations[i];
      try {
        this.applyOperation(workbook, op, i + 1);
      } catch (err) {
        // Re-throw NestJS HTTP exceptions as-is
        if (err instanceof BadRequestException || err instanceof InternalServerErrorException) {
          throw err;
        }
        this.logger.error(
          `Operation ${i + 1}/${plan.operations.length} failed: ${op.type}`,
          err.message,
        );
        throw new InternalServerErrorException(
          `Operation ${i + 1} (${op.type}) failed: ${err.message}`,
        );
      }
    }

    this.logger.log(
      `All ${plan.operations.length} operation(s) applied successfully`,
    );
  }

  // ─── Private: dispatcher ─────────────────────────────────────────────────

  private applyOperation(
    workbook: ExcelJS.Workbook,
    op: ExcelOperation,
    index: number,
  ): void {
    const worksheet = workbook.getWorksheet(op.sheet);
    if (!worksheet) {
      throw new BadRequestException(
        `Operation ${index} (${op.type}): sheet "${op.sheet}" not found in workbook. ` +
          `Available sheets: ${workbook.worksheets.map((ws) => ws.name).join(', ')}`,
      );
    }

    this.logger.debug(`  [${index}] ${op.type} on sheet "${op.sheet}"`);

    switch (op.type) {
      case 'SET_CELL':
        return this.applySetCell(worksheet, op.params, index);
      case 'SET_CELL_RANGE':
        return this.applySetCellRange(worksheet, op.params, index);
      case 'ADD_ROW':
        return this.applyAddRow(worksheet, op.params);
      case 'DELETE_ROW':
        return this.applyDeleteRow(worksheet, op.params, index);
      case 'ADD_FORMULA':
        return this.applyAddFormula(worksheet, op.params, index);
      case 'FORMAT_CELL':
        return this.applyFormatCell(worksheet, op.params, index);
      case 'ADD_COLUMN_HEADER':
        return this.applyAddColumnHeader(worksheet, op.params, index);
      case 'SORT_SHEET':
        return this.applySortSheet(worksheet, op.params, index);
      default:
        // Should never reach here — OperationsService rejects unknown types
        throw new InternalServerErrorException(
          `Unhandled operation type: ${op.type}`,
        );
    }
  }

  // ─── Private: operation implementations ──────────────────────────────────

  /** SET_CELL: set a single cell value. params: { row, col, value, allowFormulaOverwrite? } */
  private applySetCell(
    ws: ExcelJS.Worksheet,
    params: Record<string, any>,
    index: number,
  ): void {
    const row = Number(params.row);
    const col = Number(params.col);

    if (!Number.isInteger(row) || row < 1) {
      throw new BadRequestException(
        `Operation ${index} SET_CELL: row must be a positive integer, got ${params.row}`,
      );
    }
    if (!Number.isInteger(col) || col < 1) {
      throw new BadRequestException(
        `Operation ${index} SET_CELL: col must be a positive integer, got ${params.col}`,
      );
    }

    const cell = ws.getCell(row, col);

    // ── Formula overwrite protection (execution layer) ─────────────────────
    // OperationsService validates the flag at plan-validation time.
    // This second check at execution time is a defence-in-depth measure:
    // it prevents a future code path from bypassing the protection
    // even if validation is somehow skipped or a new call site is added.
    if (cell.type === ExcelJS.ValueType.Formula) {
      const allowed = params.allowFormulaOverwrite === true;
      if (!allowed) {
        const formulaValue = cell.value as ExcelJS.CellFormulaValue;
        const formula = formulaValue?.formula ?? '(unknown)';
        const cached = formulaValue?.result ?? null;
        throw new BadRequestException(
          `Operation ${index} SET_CELL: cell ${cell.address} contains a formula ` +
            `"=${formula}" (cached result: ${cached}). ` +
            `Overwriting a formula is not allowed by default. ` +
            `Set allowFormulaOverwrite: true in the operation params to explicitly authorize this.`,
        );
      }
    }

    cell.value = params.value ?? null;
  }

  /** SET_CELL_RANGE: set same value across a range. params: { startRow, startCol, endRow, endCol, value, allowFormulaOverwrite? } */
  private applySetCellRange(
    ws: ExcelJS.Worksheet,
    params: Record<string, any>,
    index: number,
  ): void {
    const startRow = Number(params.startRow);
    const startCol = Number(params.startCol);
    const endRow = Number(params.endRow);
    const endCol = Number(params.endCol);

    if (startRow < 1 || startCol < 1 || endRow < startRow || endCol < startCol) {
      throw new BadRequestException(
        `Operation ${index} SET_CELL_RANGE: invalid range ` +
          `(${startRow},${startCol}) → (${endRow},${endCol})`,
      );
    }

    // ── Formula overwrite protection — pre-scan (execution layer) ────────────
    //
    // Atomicity guarantee: scan the ENTIRE range for formula cells BEFORE
    // writing any cell. This prevents partial writes where some cells are
    // updated before a protected formula is discovered mid-range.
    //
    // If allowFormulaOverwrite is not explicitly true, collect all formula
    // cells found in the range and throw a single descriptive error.
    const allowed = params.allowFormulaOverwrite === true;

    if (!allowed) {
      const protectedCells: Array<{ address: string; formula: string; cached: string | null }> = [];

      for (let r = startRow; r <= endRow; r++) {
        for (let c = startCol; c <= endCol; c++) {
          const cell = ws.getCell(r, c);
          if (cell.type === ExcelJS.ValueType.Formula) {
            const formulaValue = cell.value as ExcelJS.CellFormulaValue;
            const rawCached = formulaValue?.result;
            const cached =
              rawCached !== undefined && rawCached !== null
                ? String(rawCached)
                : null;
            protectedCells.push({
              address: cell.address,
              formula: formulaValue?.formula ?? '(unknown)',
              cached,
            });
          }
        }
      }

      if (protectedCells.length > 0) {
        const cellList = protectedCells
          .map((c) => `${c.address} (="${ c.formula}", cached: ${c.cached})`)
          .join('; ');
        throw new BadRequestException(
          `Operation ${index} SET_CELL_RANGE: the range contains ${protectedCells.length} formula cell(s) ` +
            `that cannot be overwritten by default: ${cellList}. ` +
            `Set allowFormulaOverwrite: true in the operation params to explicitly authorize this. ` +
            `No cells in the range were modified.`,
        );
      }
    }

    // ── Write phase — only reached when no protected formulas block the op ───
    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        ws.getCell(r, c).value = params.value ?? null;
      }
    }
  }

  /** ADD_ROW: append values as a new row. params: { values: any[] } */
  private applyAddRow(
    ws: ExcelJS.Worksheet,
    params: Record<string, any>,
  ): void {
    ws.addRow(params.values as any[]);
  }

  /** DELETE_ROW: remove a row by 1-based row number. params: { row } */
  private applyDeleteRow(
    ws: ExcelJS.Worksheet,
    params: Record<string, any>,
    index: number,
  ): void {
    const row = Number(params.row);

    if (!Number.isInteger(row) || row < 1) {
      throw new BadRequestException(
        `Operation ${index} DELETE_ROW: row must be a positive integer, got ${params.row}`,
      );
    }
    if (row > ws.rowCount) {
      throw new BadRequestException(
        `Operation ${index} DELETE_ROW: row ${row} does not exist ` +
          `(sheet has ${ws.rowCount} rows)`,
      );
    }

    // spliceRows(start, deleteCount) — removes rows and shifts remaining rows up
    ws.spliceRows(row, 1);
  }

  /** ADD_FORMULA: set an Excel formula in a cell. params: { row, col, formula } */
  private applyAddFormula(
    ws: ExcelJS.Worksheet,
    params: Record<string, any>,
    index: number,
  ): void {
    const row = Number(params.row);
    const col = Number(params.col);

    if (!Number.isInteger(row) || row < 1) {
      throw new BadRequestException(
        `Operation ${index} ADD_FORMULA: row must be a positive integer`,
      );
    }
    if (!Number.isInteger(col) || col < 1) {
      throw new BadRequestException(
        `Operation ${index} ADD_FORMULA: col must be a positive integer`,
      );
    }
    if (typeof params.formula !== 'string' || !params.formula.trim()) {
      throw new BadRequestException(
        `Operation ${index} ADD_FORMULA: formula must be a non-empty string`,
      );
    }

    // ExcelJS formula cell format
    ws.getCell(row, col).value = { formula: params.formula.trim() };
  }

  /** FORMAT_CELL: apply ExcelJS style properties. params: { row, col, style: { font?, fill?, alignment?, border? } } */
  private applyFormatCell(
    ws: ExcelJS.Worksheet,
    params: Record<string, any>,
    index: number,
  ): void {
    const row = Number(params.row);
    const col = Number(params.col);

    if (!Number.isInteger(row) || row < 1) {
      throw new BadRequestException(
        `Operation ${index} FORMAT_CELL: row must be a positive integer`,
      );
    }
    if (!Number.isInteger(col) || col < 1) {
      throw new BadRequestException(
        `Operation ${index} FORMAT_CELL: col must be a positive integer`,
      );
    }
    if (!params.style || typeof params.style !== 'object') {
      throw new BadRequestException(
        `Operation ${index} FORMAT_CELL: style must be an object`,
      );
    }

    const cell = ws.getCell(row, col);
    const style = params.style as Record<string, any>;

    // Apply only supported style properties — ignore unknown keys safely
    if (style.font) cell.font = style.font;
    if (style.fill) cell.fill = style.fill;
    if (style.alignment) cell.alignment = style.alignment;
    if (style.border) cell.border = style.border;
    if (style.numFmt) cell.numFmt = style.numFmt;
  }

  /** ADD_COLUMN_HEADER: set a header value in row 1. params: { col, header } */
  private applyAddColumnHeader(
    ws: ExcelJS.Worksheet,
    params: Record<string, any>,
    index: number,
  ): void {
    const col = Number(params.col);

    if (!Number.isInteger(col) || col < 1) {
      throw new BadRequestException(
        `Operation ${index} ADD_COLUMN_HEADER: col must be a positive integer`,
      );
    }
    if (typeof params.header !== 'string') {
      throw new BadRequestException(
        `Operation ${index} ADD_COLUMN_HEADER: header must be a string`,
      );
    }

    ws.getCell(1, col).value = params.header;
  }

  /**
   * SORT_SHEET: sort data rows (row 2 onwards) by a column.
   * params: { col, order: 'asc' | 'desc' }
   *
   * ExcelJS has no built-in sort — we read, sort, and rewrite data rows.
   * Row 1 (headers) is always preserved.
   */
  private applySortSheet(
    ws: ExcelJS.Worksheet,
    params: Record<string, any>,
    index: number,
  ): void {
    const col = Number(params.col);
    const order = params.order === 'desc' ? 'desc' : 'asc';

    if (!Number.isInteger(col) || col < 1) {
      throw new BadRequestException(
        `Operation ${index} SORT_SHEET: col must be a positive integer`,
      );
    }

    if (ws.rowCount < 2) return; // nothing to sort

    // Read all data rows (2 onwards) into memory
    // row.values is 1-indexed with undefined at position 0 — we slice it off
    const dataRows: any[][] = [];
    for (let ri = 2; ri <= ws.rowCount; ri++) {
      const row = ws.getRow(ri);
      dataRows.push((Array.from(row.values as any[]) as any[]).slice(1));
    }

    const colIndex = col - 1; // convert 1-based to 0-based array index

    dataRows.sort((a, b) => {
      const av = a[colIndex];
      const bv = b[colIndex];

      // Numeric comparison
      if (typeof av === 'number' && typeof bv === 'number') {
        return order === 'asc' ? av - bv : bv - av;
      }
      // String comparison
      const as = av === null || av === undefined ? '' : String(av);
      const bs = bv === null || bv === undefined ? '' : String(bv);
      const cmp = as.localeCompare(bs);
      return order === 'asc' ? cmp : -cmp;
    });

    // Rewrite the data rows back into the worksheet
    dataRows.forEach((rowData, i) => {
      const rowNum = i + 2;
      rowData.forEach((val, ci) => {
        ws.getCell(rowNum, ci + 1).value = val;
      });
    });
  }

  // ─── Private: inspection helpers ─────────────────────────────────────────

  private inspectSheet(worksheet: ExcelJS.Worksheet): SheetInfo {
    const headers: string[] = [];

    if (worksheet.actualRowCount > 0) {
      const firstRow = worksheet.getRow(1);
      firstRow.eachCell({ includeEmpty: true }, (cell) => {
        headers.push(this.cellValueToString(cell.value));
      });
    }

    return {
      name: worksheet.name,
      rowCount: worksheet.rowCount,
      dataRowCount: worksheet.actualRowCount,
      columnCount: worksheet.columnCount,
      dataColumnCount: worksheet.actualColumnCount,
      headers,
    };
  }

  private cellValueToString(value: ExcelJS.CellValue): string {
    if (value === null || value === undefined) return '';

    if (typeof value === 'object' && 'formula' in value) {
      const result = (value as ExcelJS.CellFormulaValue).result;
      return result !== undefined && result !== null ? String(result) : '';
    }
    if (typeof value === 'object' && 'richText' in value) {
      return (value as ExcelJS.CellRichTextValue).richText
        .map((rt) => rt.text)
        .join('');
    }
    if (typeof value === 'object' && 'text' in value && 'hyperlink' in value) {
      return String((value as ExcelJS.CellHyperlinkValue).text);
    }
    if (value instanceof Date) {
      return value.toISOString();
    }
    return String(value);
  }
}
