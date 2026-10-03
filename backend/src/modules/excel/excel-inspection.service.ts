/**
 * excel-inspection.service.ts
 *
 * Phase 1 — Deep Workbook Inspection Service.
 *
 * READ-ONLY. This service never modifies workbooks.
 *
 * Provides:
 *   searchWorkbook()   — deterministic text/value search across sheets
 *   readCellRange()    — bounded structured range reader
 *   inspectCell()      — deep single-cell inspection with merge awareness
 *
 * These methods are the future evidence API for the anti-hallucination agent:
 * a cell location is only valid if it was returned by one of these methods
 * during the current agent execution.
 *
 * Uses the existing loadWorkbook() logic from ExcelService where needed.
 */

import { Injectable, Logger, BadRequestException } from '@nestjs/common';
import * as ExcelJS from 'exceljs';
import { parseCell, parseRange } from './utils/coordinate.utils';
import {
  InspectedCell,
  InspectedCellType,
  InspectedRange,
  MergeInfo,
  MergeStatus,
  RangeReaderOptions,
  SearchMatch,
  SearchOptions,
  SearchResult,
  DEFAULT_SEARCH_OPTS,
  DEFAULT_MAX_RANGE_CELLS,
} from './inspection.types';
import { WorkbookStructureService } from './workbook-structure.service';

@Injectable()
export class ExcelInspectionService {
  private readonly logger = new Logger(ExcelInspectionService.name);

  constructor(
    private readonly structureService: WorkbookStructureService,
  ) {}

  // ─── Public API ────────────────────────────────────────────────────────────

  /**
   * Search all (or a specific) worksheet for cells matching the query.
   *
   * - Searches string and numeric cell values.
   * - Case-insensitive by default.
   * - Deterministic — no LLM involvement.
   * - Bounded by maxMatches and maxCellsScanned.
   *
   * Returns: structured SearchResult with all matches and scan metadata.
   */
  async searchWorkbook(
    buffer: Buffer,
    query: string,
    options?: SearchOptions,
  ): Promise<SearchResult> {
    if (!query || typeof query !== 'string' || !query.trim()) {
      throw new BadRequestException(
        'Search query must be a non-empty string',
      );
    }

    const opts = this.resolveSearchOptions(options);
    const workbook = await this.load(buffer);

    const matches: SearchMatch[] = [];
    let cellsScanned = 0;
    let truncated = false;

    const queryNorm = opts.caseInsensitive
      ? query.trim().toLowerCase()
      : query.trim();

    for (const worksheet of workbook.worksheets) {
      // Apply optional sheet filter
      if (opts.sheetName && worksheet.name !== opts.sheetName) {
        continue;
      }

      const sheetDone = this.searchSheet(
        worksheet,
        queryNorm,
        opts,
        matches,
        cellsScanned,
      );

      cellsScanned = sheetDone.cellsScanned;
      truncated = sheetDone.truncated;

      if (truncated) break;
    }

    // Optionally enrich each match with surrounding context evidence
    if (opts.includeContext && matches.length > 0) {
      // We need to reload worksheets for context building.
      // Re-use the already-loaded workbook by iterating worksheets by name.
      const contextWorkbook = await this.load(buffer);
      for (const match of matches) {
        const ws = contextWorkbook.getWorksheet(match.sheetName);
        if (ws) {
          try {
            match.context = await this.structureService.buildSearchMatchContext(
              ws,
              match.sheetName,
              match.coordinate,
            );
          } catch {
            // Context enrichment is best-effort — never block the search result
          }
        }
      }
    }

    this.logger.log(
      `Search "${query}" → ${matches.length} match(es), ${cellsScanned} cells scanned, ` +
        `truncated=${truncated}, context=${opts.includeContext}`,
    );

    return {
      query,
      options: opts,
      matches,
      cellsScanned,
      truncated,
    };
  }

  /**
   * Read a bounded cell range from a named worksheet.
   *
   * Returns: structured range data including value, formula, type, merge info.
   *
   * Throws if:
   *   - worksheet not found
   *   - range is invalid
   *   - range exceeds maxCells limit
   */
  async readCellRange(
    buffer: Buffer,
    sheetName: string,
    range: string,
    options?: RangeReaderOptions,
  ): Promise<InspectedRange> {
    this.validateSheetNameInput(sheetName);

    const maxCells = options?.maxCells ?? DEFAULT_MAX_RANGE_CELLS;

    // Validate and parse range
    let parsed;
    try {
      parsed = parseRange(range);
    } catch (err) {
      throw new BadRequestException(
        `Invalid range "${range}": ${err.message}`,
      );
    }

    if (parsed.rows * parsed.cols > maxCells) {
      throw new BadRequestException(
        `Range "${range}" contains ${parsed.rows * parsed.cols} cells, ` +
          `which exceeds the limit of ${maxCells}. Use a smaller range.`,
      );
    }

    const workbook = await this.load(buffer);
    const worksheet = workbook.getWorksheet(sheetName);

    if (!worksheet) {
      throw new BadRequestException(
        `Worksheet "${sheetName}" not found. ` +
          `Available: ${workbook.worksheets.map((w) => `"${w.name}"`).join(', ')}`,
      );
    }

    const cells: InspectedCell[] = [];

    for (let r = parsed.start.row; r <= parsed.end.row; r++) {
      for (let c = parsed.start.col; c <= parsed.end.col; c++) {
        const exCell = worksheet.getCell(r, c);
        cells.push(
          this.buildInspectedCell(sheetName, exCell),
        );
      }
    }

    this.logger.log(
      `readCellRange "${sheetName}" ${parsed.address} → ${cells.length} cell(s)`,
    );

    return {
      sheetName,
      range: parsed.address,
      dimensions: {
        rows: parsed.rows,
        cols: parsed.cols,
        totalCells: cells.length,
      },
      cells,
    };
  }

  /**
   * Deep inspection of a single cell.
   *
   * Returns: value, formula, type, numFmt, merge membership.
   *
   * This is the primary evidence API:
   * before the future agent modifies G42, it must call inspectCell("Sheet", "G42")
   * and verify the returned evidence matches the intended target.
   *
   * Throws if:
   *   - worksheet not found
   *   - coordinate is invalid
   */
  /**
   * Find all cells in the workbook whose formulas reference any of the given
   * target coordinates. Used for cross-sheet dependency detection after a write plan
   * is produced — so the agent can ask the user before silently missing dependents.
   *
   * Strategy: string-match each formula against the target coordinate patterns.
   * Covers the common 80% case (direct references like B3, Sheet2!B3, $B$3).
   * Does NOT resolve INDIRECT/OFFSET/dynamic references — those are uncommon.
   *
   * @param buffer      The workbook buffer (read-only)
   * @param targets     Array of { sheetName, coordinate } pairs being written
   * @returns           Array of dependent cells found
   */
  async findDependentCells(
    buffer: Buffer,
    targets: Array<{ sheetName: string; coordinate: string }>,
  ): Promise<Array<{ sheetName: string; coordinate: string; formula: string; dependsOn: string }>> {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer as any);

    const dependents: Array<{
      sheetName: string;
      coordinate: string;
      formula: string;
      dependsOn: string;
    }> = [];

    // Build reference patterns for each target
    // e.g. target B3 on "Fund Details" → ['B3', '$B$3', '$B3', 'B$3', "'Fund Details'!B3", etc.]
    const patterns: Array<{ pattern: RegExp; label: string }> = [];
    for (const t of targets) {
      const col = t.coordinate.replace(/[0-9]/g, '');
      const row = t.coordinate.replace(/[^0-9]/g, '');
      // Basic ref: B3, $B$3, $B3, B$3
      const basicRef = `\\$?${col}\\$?${row}(?![0-9A-Z])`;
      // Sheet-qualified: 'Fund Details'!B3 or Fund_Details!B3
      const sheetRef = `'?${t.sheetName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}'?!\\$?${col}\\$?${row}`;
      patterns.push({
        pattern: new RegExp(`(${sheetRef}|${basicRef})`, 'i'),
        label: `${t.sheetName}!${t.coordinate}`,
      });
    }

    wb.eachSheet(ws => {
      ws.eachRow(row => {
        row.eachCell(cell => {
          if (!cell.formula) return;
          for (const { pattern, label } of patterns) {
            if (pattern.test(cell.formula)) {
              // Don't report the target cell itself
              const cellCoord = cell.address;
              const alreadyTarget = targets.some(
                t => t.sheetName === ws.name && t.coordinate === cellCoord,
              );
              if (!alreadyTarget) {
                dependents.push({
                  sheetName: ws.name,
                  coordinate: cellCoord,
                  formula: cell.formula,
                  dependsOn: label,
                });
              }
            }
          }
        });
      });
    });

    return dependents;
  }

  async inspectCell(
    buffer: Buffer,
    sheetName: string,
    coordinate: string,
  ): Promise<InspectedCell> {
    this.validateSheetNameInput(sheetName);

    let parsed;
    try {
      parsed = parseCell(coordinate);
    } catch (err) {
      throw new BadRequestException(
        `Invalid coordinate "${coordinate}": ${err.message}`,
      );
    }

    const workbook = await this.load(buffer);
    const worksheet = workbook.getWorksheet(sheetName);

    if (!worksheet) {
      throw new BadRequestException(
        `Worksheet "${sheetName}" not found. ` +
          `Available: ${workbook.worksheets.map((w) => `"${w.name}"`).join(', ')}`,
      );
    }

    const exCell = worksheet.getCell(parsed.row, parsed.col);
    const result = this.buildInspectedCell(sheetName, exCell);

    this.logger.log(
      `inspectCell "${sheetName}" ${parsed.address} → type=${result.type}, ` +
        `formula=${result.formula ? 'yes' : 'no'}, merge=${result.merge.status}`,
    );

    return result;
  }

  // ─── Private: workbook loading ────────────────────────────────────────────

  /**
   * Load a workbook from buffer.
   * Reuses the same pattern as ExcelService.loadWorkbook() to avoid duplication.
   * Kept private here — callers always use the public inspection methods.
   */
  private async load(buffer: Buffer): Promise<ExcelJS.Workbook> {
    if (!buffer || buffer.length === 0) {
      throw new BadRequestException('File buffer is empty');
    }
    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(
        buffer as unknown as Parameters<typeof workbook.xlsx.load>[0],
      );
      return workbook;
    } catch (err) {
      throw new BadRequestException(
        `Could not read Excel file: ${err.message}`,
      );
    }
  }

  // ─── Private: search ─────────────────────────────────────────────────────

  private resolveSearchOptions(options?: SearchOptions): Required<SearchOptions> {
    return {
      caseInsensitive: options?.caseInsensitive ?? DEFAULT_SEARCH_OPTS.caseInsensitive,
      sheetName: options?.sheetName ?? DEFAULT_SEARCH_OPTS.sheetName,
      maxMatches: options?.maxMatches ?? DEFAULT_SEARCH_OPTS.maxMatches,
      maxCellsScanned: options?.maxCellsScanned ?? DEFAULT_SEARCH_OPTS.maxCellsScanned,
      includeContext: options?.includeContext ?? DEFAULT_SEARCH_OPTS.includeContext,
    };
  }

  private searchSheet(
    worksheet: ExcelJS.Worksheet,
    queryNorm: string,
    opts: Required<SearchOptions>,
    matches: SearchMatch[],
    cellsScannedSoFar: number,
  ): { cellsScanned: number; truncated: boolean } {
    let cellsScanned = cellsScannedSoFar;
    let truncated = false;

    worksheet.eachRow((row, rowIndex) => {
      if (truncated) return; // already hit a limit

      row.eachCell({ includeEmpty: false }, (cell, colIndex) => {
        if (truncated) return;

        cellsScanned++;

        // Check scan limit
        if (cellsScanned > opts.maxCellsScanned) {
          truncated = true;
          return;
        }

        // Check match limit
        if (matches.length >= opts.maxMatches) {
          truncated = true;
          return;
        }

        const { displayValue, type } = this.getCellDisplayValue(cell);

        if (displayValue === null) return; // skip truly empty

        const displayStr = String(displayValue);
        const compareStr = opts.caseInsensitive
          ? displayStr.toLowerCase()
          : displayStr;

        if (compareStr.includes(queryNorm)) {
          const address = cell.address;
          matches.push({
            sheetName: worksheet.name,
            coordinate: address,
            value: displayStr,
            type,
          });
        }
      });
    });

    return { cellsScanned, truncated };
  }

  // ─── Private: cell building ───────────────────────────────────────────────

  private buildInspectedCell(
    sheetName: string,
    cell: ExcelJS.Cell,
  ): InspectedCell {
    const { displayValue, formula, type } = this.getCellDisplayValue(cell);

    const numFmt =
      cell.numFmt && cell.numFmt !== 'General' ? cell.numFmt : null;

    const merge = this.buildMergeInfo(cell);

    return {
      sheetName,
      coordinate: cell.address,
      value: displayValue,
      formula,
      type,
      numFmt,
      merge,
    };
  }

  /**
   * Extract the display value, formula, and type from an ExcelJS cell.
   *
   * ExcelJS ValueType enum:
   *   0 = Null, 1 = Merge, 2 = Number, 3 = String, 4 = Date,
   *   5 = Hyperlink, 6 = Formula, 7 = SharedString, 8 = RichText,
   *   9 = Boolean, 10 = Error
   */
  private getCellDisplayValue(cell: ExcelJS.Cell): {
    displayValue: string | number | boolean | null;
    formula: string | null;
    type: InspectedCellType;
  } {
    const vt = cell.type; // ExcelJS ValueType number

    // Empty / Null
    if (vt === 0) {
      return { displayValue: null, formula: null, type: 'empty' };
    }

    // Merge-member cell: ExcelJS type 1 means this cell is a non-master
    // member of a merged range. The actual value lives on the master cell.
    if (vt === 1) {
      // For merge members, still read the value ExcelJS exposes
      // (which is the master's value in many cases)
      const rawVal = cell.value;
      const displayValue = rawVal !== null && rawVal !== undefined
        ? this.extractRawValue(rawVal)
        : null;
      return { displayValue, formula: null, type: 'empty' };
    }

    // Formula
    if (vt === 6) {
      const formulaValue = cell.value as ExcelJS.CellFormulaValue;
      const formulaStr =
        typeof formulaValue?.formula === 'string'
          ? formulaValue.formula
          : null;
      const result = formulaValue?.result;
      const displayValue =
        result !== undefined && result !== null
          ? this.extractRawValue(result)
          : null;
      return {
        displayValue,
        formula: formulaStr,
        type: 'formula',
      };
    }

    // Number
    if (vt === 2) {
      return {
        displayValue: typeof cell.value === 'number' ? cell.value : null,
        formula: null,
        type: 'number',
      };
    }

    // Date
    if (vt === 4) {
      const d = cell.value;
      return {
        displayValue: d instanceof Date ? d.toISOString() : String(d),
        formula: null,
        type: 'date',
      };
    }

    // Boolean
    if (vt === 9) {
      return {
        displayValue: typeof cell.value === 'boolean' ? cell.value : null,
        formula: null,
        type: 'boolean',
      };
    }

    // String, SharedString, RichText, Hyperlink — all become string
    const strVal = this.extractStringValue(cell.value);
    return {
      displayValue: strVal,
      formula: null,
      type: strVal !== null ? 'string' : 'empty',
    };
  }

  /**
   * Extract a primitive value from any ExcelJS cell value type.
   */
  private extractRawValue(
    val: ExcelJS.CellValue,
  ): string | number | boolean | null {
    if (val === null || val === undefined) return null;
    if (typeof val === 'string') return val;
    if (typeof val === 'number') return val;
    if (typeof val === 'boolean') return val;
    if (val instanceof Date) return val.toISOString();

    // RichText
    if (typeof val === 'object' && 'richText' in val) {
      return (val as ExcelJS.CellRichTextValue).richText
        .map((rt) => rt.text)
        .join('');
    }

    // Hyperlink
    if (typeof val === 'object' && 'text' in val && 'hyperlink' in val) {
      return String((val as ExcelJS.CellHyperlinkValue).text);
    }

    // Formula result
    if (typeof val === 'object' && 'formula' in val) {
      const r = (val as ExcelJS.CellFormulaValue).result;
      return r !== undefined && r !== null ? this.extractRawValue(r) : null;
    }

    return String(val);
  }

  private extractStringValue(
    val: ExcelJS.CellValue,
  ): string | null {
    const raw = this.extractRawValue(val);
    if (raw === null || raw === undefined) return null;
    return String(raw);
  }

  // ─── Private: merge handling ──────────────────────────────────────────────

  /**
   * Build merge information for a cell.
   *
   * Uses ExcelJS:
   *   cell.isMerged — true if the cell is part of any merge
   *   cell.master.address — address of the top-left (master) cell
   *   worksheet._merges — internal merge table: { masterAddr: { model: { top, left, bottom, right } } }
   *
   * Note: worksheet._merges is an internal ExcelJS API. We use it because
   * ExcelJS does not expose a public mergeArea property. This is documented
   * here so it can be updated if ExcelJS changes the internal structure.
   */
  private buildMergeInfo(cell: ExcelJS.Cell): MergeInfo {
    if (!cell.isMerged) {
      return { status: 'none', mergedRange: null, masterAddress: null };
    }

    const masterAddr = cell.master?.address ?? null;
    const isMaster = masterAddr === cell.address;

    // Attempt to resolve the full merged range from the internal merge table
    const mergedRange = this.resolveMergeRange(cell, masterAddr);

    const status: MergeStatus = isMaster ? 'master' : 'member';

    return {
      status,
      mergedRange,
      masterAddress: masterAddr,
    };
  }

  private resolveMergeRange(
    cell: ExcelJS.Cell,
    masterAddr: string | null,
  ): string | null {
    try {
      // Access the worksheet through the cell's model parent
      // ExcelJS cell has a reference to its worksheet via _row._worksheet
      const ws = (cell as any)._row?._worksheet;
      if (!ws || !ws._merges) return null;

      const effectiveMaster = masterAddr ?? cell.address;

      const mergeModel = ws._merges[effectiveMaster];
      if (!mergeModel?.model) return null;

      const { top, left, bottom, right } = mergeModel.model;

      // Build range from numerical top/left/bottom/right
      const { colNumberToLetter } = require('./utils/coordinate.utils');
      const startAddr = `${colNumberToLetter(left)}${top}`;
      const endAddr = `${colNumberToLetter(right)}${bottom}`;

      return `${startAddr}:${endAddr}`;
    } catch {
      // If the internal API changes, degrade gracefully
      return null;
    }
  }

  // ─── Private: input validation ────────────────────────────────────────────

  private validateSheetNameInput(sheetName: string): void {
    if (!sheetName || typeof sheetName !== 'string' || !sheetName.trim()) {
      throw new BadRequestException(
        'Sheet name must be a non-empty string',
      );
    }
  }
}
