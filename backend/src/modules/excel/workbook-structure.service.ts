/**
 * workbook-structure.service.ts
 *
 * Phase 1 Step 2 — Workbook Intelligence & Structural Inspection.
 *
 * READ-ONLY. This service never modifies workbooks.
 *
 * Provides:
 *   inspectSheetStructure()  — sheet-level summary (bounds, counts, merges, hidden)
 *   detectHeaders()          — deterministic header row candidates
 *   detectSections()         — table/section region candidates
 *   getCellContext()         — surrounding rows/cols around a target cell
 *
 * All methods are evidence-based and deterministic.
 * No LLM is used. No Excel modifications are made.
 *
 * Architecture:
 *   ExcelService             → write/modification operations
 *   ExcelInspectionService   → cell/range/search inspection (Step 1)
 *   WorkbookStructureService → structural understanding (Step 2)  ← THIS FILE
 */

import { Injectable, Logger, BadRequestException } from '@nestjs/common';
import * as ExcelJS from 'exceljs';
import { parseCell, buildCellAddress, colNumberToLetter } from './utils/coordinate.utils';
import { InspectedCell, InspectedCellType } from './inspection.types';
import {
  CellContext,
  CellContextOptions,
  HeaderCandidate,
  HeaderConfidence,
  HeaderDetectionResult,
  SectionCandidate,
  SectionDetectionResult,
  SheetMergeRange,
  SheetStructure,
  SearchMatchContext,
  DEFAULT_CELL_CONTEXT_OPTIONS,
  MAX_CELL_CONTEXT_COL_RADIUS,
  MAX_CELL_CONTEXT_ROW_RADIUS,
  SECTION_MIN_CELLS,
  SECTION_SPLIT_GAP,
} from './structure.types';

@Injectable()
export class WorkbookStructureService {
  private readonly logger = new Logger(WorkbookStructureService.name);

  // ─── Public API ───────────────────────────────────────────────────────────

  /**
   * Produce a structural summary of a worksheet.
   *
   * Returns:
   *   - used bounds (first/last row/col with content)
   *   - row/col counts (total, populated, hidden)
   *   - cell type counts (text, numeric, formula, etc.)
   *   - all merged ranges (from public ws.model?.merges)
   *   - hidden row/col numbers
   */
  async inspectSheetStructure(
    buffer: Buffer,
    sheetName: string,
  ): Promise<SheetStructure> {
    this.validateSheetNameInput(sheetName);
    const workbook = await this.load(buffer);
    const ws = this.requireWorksheet(workbook, sheetName);

    const dims = ws.dimensions as any;
    const model = dims?.model;

    // Used bounds from ExcelJS dimensions
    const firstRow: number | null = model?.top > 0 ? model.top : null;
    const lastRow: number | null = model?.bottom > 0 ? model.bottom : null;
    const firstCol: number | null = model?.left > 0 ? model.left : null;
    const lastCol: number | null = model?.right > 0 ? model.right : null;

    const usedAddress =
      firstRow && lastRow && firstCol && lastCol
        ? `${buildCellAddress(firstRow, firstCol)}:${buildCellAddress(lastRow, lastCol)}`
        : null;

    // Hidden rows/cols
    const hiddenRows: number[] = [];
    const hiddenCols: number[] = [];
    let populatedRows = 0;

    // Use index-based scan for hidden rows: eachRow() skips rows that have
    // no cell content even if they are explicitly hidden (e.g. spacer rows).
    // ws.rowCount includes all rows up to the last used row.
    const totalRows = ws.rowCount;
    for (let i = 1; i <= totalRows; i++) {
      const row = ws.getRow(i);
      if ((row as any).hidden === true) {
        hiddenRows.push(i);
      }
      // Check if this row has any non-empty cells
      let hasContent = false;
      row.eachCell({ includeEmpty: false }, () => { hasContent = true; });
      if (hasContent) populatedRows++;
    }

    // Columns: iterate up to the last used column
    let populatedCols = 0;
    if (lastCol) {
      for (let c = 1; c <= lastCol; c++) {
        const col = ws.getColumn(c);
        if ((col as any).hidden === true) {
          hiddenCols.push(c);
        }
        // Count if column has any non-empty cells
        if (c <= (ws.actualColumnCount || 0)) populatedCols++;
      }
    }

    // Cell type counts
    const cellCounts = {
      nonEmpty: 0,
      text: 0,
      numeric: 0,
      date: 0,
      formula: 0,
      boolean: 0,
    };

    ws.eachRow((row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        cellCounts.nonEmpty++;
        switch (cell.type) {
          case 2: cellCounts.numeric++; break;   // Number
          case 3: cellCounts.text++; break;       // String
          case 4: cellCounts.date++; break;       // Date
          case 6: cellCounts.formula++; break;    // Formula
          case 9: cellCounts.boolean++; break;    // Boolean
          // SharedString(7), RichText(8) also count as text
          case 7: case 8: cellCounts.text++; break;
        }
      });
    });

    // Merged ranges from public API
    const mergedRanges = this.collectMergedRanges(ws);

    const result: SheetStructure = {
      sheetName,
      usedBounds: {
        firstRow,
        lastRow,
        firstCol,
        lastCol,
        address: usedAddress,
      },
      rowCount: {
        total: ws.rowCount,
        populated: populatedRows,
        hidden: hiddenRows.length,
      },
      colCount: {
        total: ws.columnCount,
        populated: populatedCols,
        hidden: hiddenCols.length,
      },
      cellCounts,
      mergedRanges,
      hiddenRows,
      hiddenCols,
    };

    this.logger.log(
      `inspectSheetStructure "${sheetName}" → ` +
        `${cellCounts.nonEmpty} cells, ${mergedRanges.length} merges, ` +
        `${hiddenRows.length} hidden rows, ${hiddenCols.length} hidden cols`,
    );

    return result;
  }

  /**
   * Detect candidate header rows in a worksheet.
   *
   * Uses structural evidence only — no LLM:
   *   - text density (ratio of text cells to total non-empty cells)
   *   - merged cells (merged headings are common in financial reports)
   *   - position relative to empty rows (headers often follow blank rows)
   *   - row is populated but surrounded by dense data rows
   *
   * Returns multiple candidates when evidence is ambiguous.
   * The agent must not assume the first candidate is definitively correct.
   */
  async detectHeaders(
    buffer: Buffer,
    sheetName: string,
  ): Promise<HeaderDetectionResult> {
    this.validateSheetNameInput(sheetName);
    const workbook = await this.load(buffer);
    const ws = this.requireWorksheet(workbook, sheetName);

    const candidates: HeaderCandidate[] = [];

    // Build a row summary for the entire used area
    const rowSummaries = this.buildRowSummaries(ws);

    for (let i = 0; i < rowSummaries.length; i++) {
      const summary = rowSummaries[i];

      // Skip completely empty rows
      if (summary.populatedCellCount === 0) continue;

      const confidence = this.scoreHeaderRow(rowSummaries, i);

      // Include rows with at least low confidence
      if (confidence !== null) {
        candidates.push({
          rowNumber: summary.rowNumber,
          coordinates: summary.coordinates,
          labels: summary.values,
          populatedCellCount: summary.populatedCellCount,
          textCellCount: summary.textCellCount,
          hasMergedCells: summary.hasMergedCells,
          confidence,
          evidence: this.buildHeaderEvidence(rowSummaries, i, confidence),
        });
      }
    }

    // Sort: high > medium > low confidence, then by row number
    const confidenceOrder: Record<HeaderConfidence, number> = {
      high: 0,
      medium: 1,
      low: 2,
    };
    candidates.sort(
      (a, b) =>
        confidenceOrder[a.confidence] - confidenceOrder[b.confidence] ||
        a.rowNumber - b.rowNumber,
    );

    this.logger.log(
      `detectHeaders "${sheetName}" → ${candidates.length} candidate(s)`,
    );

    return {
      sheetName,
      candidates,
      hasMultipleCandidates: candidates.length > 1,
    };
  }

  /**
   * Detect candidate table/section regions in a worksheet.
   *
   * A section is a contiguous group of populated rows separated from
   * others by SECTION_SPLIT_GAP or more consecutive empty rows.
   *
   * Returns sections ordered by their start row.
   * Each section includes its own header candidates.
   */
  async detectSections(
    buffer: Buffer,
    sheetName: string,
  ): Promise<SectionDetectionResult> {
    this.validateSheetNameInput(sheetName);
    const workbook = await this.load(buffer);
    const ws = this.requireWorksheet(workbook, sheetName);

    const rowSummaries = this.buildRowSummaries(ws);
    const sections: SectionCandidate[] = [];

    let sectionStart: number | null = null;
    let consecutiveEmpty = 0;
    let currentSectionRows: number[] = [];

    const flushSection = () => {
      if (currentSectionRows.length === 0) return;

      const allRows = currentSectionRows;
      const firstRowNum = allRows[0];
      const lastRowNum = allRows[allRows.length - 1];

      // Count non-empty cells and detect formulas/merges
      let nonEmptyCells = 0;
      let hasFormulas = false;
      let hasMergedCells = false;
      let minCol = Infinity;
      let maxCol = -Infinity;

      for (const rowNum of allRows) {
        const summary = rowSummaries.find((r) => r.rowNumber === rowNum);
        if (!summary) continue;
        nonEmptyCells += summary.populatedCellCount;
        if (summary.hasFormulas) hasFormulas = true;
        if (summary.hasMergedCells) hasMergedCells = true;
        if (summary.minCol < minCol) minCol = summary.minCol;
        if (summary.maxCol > maxCol) maxCol = summary.maxCol;
      }

      if (nonEmptyCells < SECTION_MIN_CELLS) return;

      if (minCol === Infinity) minCol = 1;
      if (maxCol === -Infinity) maxCol = 1;

      const startCoord = buildCellAddress(firstRowNum, minCol);
      const endCoord = buildCellAddress(lastRowNum, maxCol);
      const rangeAddress = `${startCoord}:${endCoord}`;

      // Find header candidates within/above this section
      const headerResult = this.scoreHeaderCandidatesForSection(
        rowSummaries,
        firstRowNum,
        lastRowNum,
      );

      sections.push({
        index: sections.length,
        startCoordinate: startCoord,
        endCoordinate: endCoord,
        rangeAddress,
        rows: lastRowNum - firstRowNum + 1,
        cols: maxCol - minCol + 1,
        nonEmptyCells,
        hasFormulas,
        hasMergedCells,
        headerCandidates: headerResult,
      });

      currentSectionRows = [];
      sectionStart = null;
    };

    for (const summary of rowSummaries) {
      if (summary.populatedCellCount > 0) {
        // Non-empty row — part of a section
        consecutiveEmpty = 0;
        currentSectionRows.push(summary.rowNumber);
      } else {
        // Empty row
        consecutiveEmpty++;
        if (consecutiveEmpty >= SECTION_SPLIT_GAP && currentSectionRows.length > 0) {
          flushSection();
          consecutiveEmpty = 0;
        }
      }
    }
    // Flush final section
    flushSection();

    this.logger.log(
      `detectSections "${sheetName}" → ${sections.length} section(s)`,
    );

    return { sheetName, sections };
  }

  /**
   * Inspect the surrounding context of a target cell.
   *
   * Returns the target cell plus nearby cells within a configurable window
   * (rowRadius rows above/below, colRadius cols left/right).
   *
   * Also returns row labels (other non-empty cells in the same row) and
   * column labels (non-empty cells in the same column) to help the agent
   * understand what the target cell represents.
   */
  async getCellContext(
    buffer: Buffer,
    sheetName: string,
    coordinate: string,
    options?: CellContextOptions,
  ): Promise<CellContext> {
    this.validateSheetNameInput(sheetName);

    let parsed: ReturnType<typeof parseCell>;
    try {
      parsed = parseCell(coordinate);
    } catch (err) {
      throw new BadRequestException(
        `Invalid coordinate "${coordinate}": ${err.message}`,
      );
    }

    const rowRadius = Math.min(
      options?.rowRadius ?? DEFAULT_CELL_CONTEXT_OPTIONS.rowRadius,
      MAX_CELL_CONTEXT_ROW_RADIUS,
    );
    const colRadius = Math.min(
      options?.colRadius ?? DEFAULT_CELL_CONTEXT_OPTIONS.colRadius,
      MAX_CELL_CONTEXT_COL_RADIUS,
    );

    const workbook = await this.load(buffer);
    const ws = this.requireWorksheet(workbook, sheetName);

    // Clamp context window to actual sheet bounds
    const dims = ws.dimensions as any;
    const model = dims?.model;
    const sheetFirstRow = model?.top ?? 1;
    const sheetLastRow = model?.bottom ?? ws.rowCount;
    const sheetFirstCol = model?.left ?? 1;
    const sheetLastCol = model?.right ?? ws.columnCount;

    const startRow = Math.max(sheetFirstRow, parsed.row - rowRadius);
    const endRow = Math.min(sheetLastRow, parsed.row + rowRadius);
    const startCol = Math.max(sheetFirstCol, parsed.col - colRadius);
    const endCol = Math.min(sheetLastCol, parsed.col + colRadius);

    const contextAddress = `${buildCellAddress(startRow, startCol)}:${buildCellAddress(endRow, endCol)}`;

    // Read all context cells
    const cells: InspectedCell[] = [];
    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        cells.push(this.buildInspectedCell(sheetName, ws.getCell(r, c)));
      }
    }

    // Target cell
    const targetCell = this.buildInspectedCell(
      sheetName,
      ws.getCell(parsed.row, parsed.col),
    );

    // Row labels: other non-empty cells in the same row
    const rowLabels: InspectedCell[] = [];
    if (sheetLastCol) {
      for (let c = sheetFirstCol; c <= sheetLastCol; c++) {
        if (c === parsed.col) continue;
        const cell = ws.getCell(parsed.row, c);
        const inspected = this.buildInspectedCell(sheetName, cell);
        if (inspected.type !== 'empty') {
          rowLabels.push(inspected);
        }
      }
    }

    // Column labels: non-empty cells in the same column (limited to context window)
    const colLabels: InspectedCell[] = [];
    for (let r = startRow; r <= endRow; r++) {
      if (r === parsed.row) continue;
      const cell = ws.getCell(r, parsed.col);
      const inspected = this.buildInspectedCell(sheetName, cell);
      if (inspected.type !== 'empty') {
        colLabels.push(inspected);
      }
    }

    this.logger.log(
      `getCellContext "${sheetName}" ${parsed.address} ` +
        `± (${rowRadius}r, ${colRadius}c) → ${cells.length} context cells`,
    );

    return {
      sheetName,
      targetCell,
      contextRange: {
        startRow,
        endRow,
        startCol,
        endCol,
        address: contextAddress,
      },
      cells,
      rowLabels,
      colLabels,
    };
  }

  /**
   * Build enhanced context for a specific search match coordinate.
   *
   * Used internally to enrich SearchMatch results when requested.
   * Returns surrounding row labels and nearby headings.
   */
  async buildSearchMatchContext(
    ws: ExcelJS.Worksheet,
    sheetName: string,
    coordinate: string,
  ): Promise<SearchMatchContext> {
    const parsed = parseCell(coordinate);
    const cell = ws.getCell(parsed.row, parsed.col);
    const inspectedCell = this.buildInspectedCell(sheetName, cell);

    // Row labels: other non-empty cells in the same row
    const rowLabels: SearchMatchContext['rowLabels'] = [];
    const dims = ws.dimensions as any;
    const model = dims?.model;
    const lastCol = model?.right ?? ws.columnCount;

    for (let c = 1; c <= (lastCol || 1); c++) {
      if (c === parsed.col) continue;
      const rowCell = ws.getCell(parsed.row, c);
      const inspected = this.buildInspectedCell(sheetName, rowCell);
      if (inspected.type !== 'empty') {
        rowLabels.push({
          coordinate: inspected.coordinate,
          value: inspected.value,
          type: inspected.type,
        });
      }
    }

    // Nearby above: up to 2 rows above, non-empty cells
    const nearbyAbove: SearchMatchContext['nearbyAbove'] = [];
    for (let offset = 1; offset <= 2; offset++) {
      const r = parsed.row - offset;
      if (r < 1) break;
      for (let c = 1; c <= (lastCol || 1); c++) {
        const aboveCell = ws.getCell(r, c);
        const inspected = this.buildInspectedCell(sheetName, aboveCell);
        if (inspected.type !== 'empty') {
          nearbyAbove.push({
            coordinate: inspected.coordinate,
            value: inspected.value,
            rowNumber: r,
          });
        }
      }
    }

    return { cell: inspectedCell, rowLabels, nearbyAbove };
  }

  // ─── Private: workbook loading ────────────────────────────────────────────

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

  private requireWorksheet(
    workbook: ExcelJS.Workbook,
    sheetName: string,
  ): ExcelJS.Worksheet {
    const ws = workbook.getWorksheet(sheetName);
    if (!ws) {
      const available = workbook.worksheets.map((w) => `"${w.name}"`).join(', ');
      throw new BadRequestException(
        `Worksheet "${sheetName}" not found. Available: ${available}`,
      );
    }
    return ws;
  }

  // ─── Private: merge collection ────────────────────────────────────────────

  private collectMergedRanges(ws: ExcelJS.Worksheet): SheetMergeRange[] {
    const result: SheetMergeRange[] = [];

    // Use public ws.model?.merges[] API — returns array of "A1:C1" strings
    const merges: string[] = (ws as any).model?.merges ?? [];

    for (const rangeStr of merges) {
      const colonIdx = rangeStr.indexOf(':');
      if (colonIdx === -1) continue;

      const masterAddr = rangeStr.slice(0, colonIdx);
      const endAddr = rangeStr.slice(colonIdx + 1);

      try {
        const master = parseCell(masterAddr);
        const end = parseCell(endAddr);
        result.push({
          range: rangeStr,
          masterAddress: masterAddr,
          rowSpan: end.row - master.row + 1,
          colSpan: end.col - master.col + 1,
        });
      } catch {
        // Skip malformed merge entries
      }
    }

    return result;
  }

  // ─── Private: row summary ─────────────────────────────────────────────────

  /**
   * Internal per-row summary used by header and section detection.
   */
  private buildRowSummaries(ws: ExcelJS.Worksheet): RowSummary[] {
    const summaries: RowSummary[] = [];
    const dims = ws.dimensions as any;
    const model = dims?.model;
    const lastRow = model?.bottom ?? ws.rowCount;

    if (!lastRow || lastRow === 0) return summaries;

    // Use index-based iteration (not eachRow) so we see every row in the
    // used range — including hidden or empty spacer rows that eachRow skips.
    // Empty spacer rows are essential for section boundary detection.
    for (let rowNum = 1; rowNum <= lastRow; rowNum++) {
      const row = ws.getRow(rowNum);
      const coordinates: string[] = [];
      const values: (string | number | boolean | null)[] = [];
      let populatedCellCount = 0;
      let textCellCount = 0;
      let hasFormulas = false;
      let hasMergedCells = false;
      let minCol = Infinity;
      let maxCol = -Infinity;

      row.eachCell({ includeEmpty: false }, (cell, colNum) => {
        populatedCellCount++;
        coordinates.push(cell.address);
        const { displayValue, type } = this.getCellDisplayValue(cell);
        values.push(displayValue);

        if (type === 'string') textCellCount++;
        if (type === 'formula') hasFormulas = true;
        if (cell.isMerged) hasMergedCells = true;
        if (colNum < minCol) minCol = colNum;
        if (colNum > maxCol) maxCol = colNum;
      });

      summaries.push({
        rowNumber: rowNum,
        populatedCellCount,
        textCellCount,
        hasFormulas,
        hasMergedCells,
        coordinates,
        values,
        minCol: minCol === Infinity ? 0 : minCol,
        maxCol: maxCol === -Infinity ? 0 : maxCol,
        isHidden: (row as any).hidden === true,
      });
    }

    return summaries;
  }

  // ─── Private: header scoring ──────────────────────────────────────────────

  private scoreHeaderRow(
    summaries: RowSummary[],
    index: number,
  ): HeaderConfidence | null {
    const row = summaries[index];
    if (row.populatedCellCount === 0) return null;

    let score = 0;

    // Strong signal: row is mostly text
    const textRatio = row.textCellCount / row.populatedCellCount;
    if (textRatio >= 0.8) score += 3;
    else if (textRatio >= 0.5) score += 1;
    else if (textRatio === 0) return null; // All numeric rows are not headers

    // Very strong signal: merged cells in this row (financial report headings)
    if (row.hasMergedCells) score += 3;

    // Signal: preceded by an empty row or is the first row
    const prev = summaries[index - 1];
    if (!prev || prev.populatedCellCount === 0) score += 2;

    // Signal: followed by a populated row
    const next = summaries[index + 1];
    if (next && next.populatedCellCount > 0) score += 1;

    // Signal: followed by numeric/formula rows (typical for data below a header)
    if (
      next &&
      next.textCellCount < next.populatedCellCount &&
      next.populatedCellCount > 0
    ) {
      score += 1;
    }

    // Penalty: row is deeply inside a run of similarly structured rows
    // (likely a data row, not a header) — but NOT if it has merged cells
    if (!row.hasMergedCells && prev && prev.populatedCellCount > 0 && index >= 2) {
      const prevPrev = summaries[index - 2];
      if (prevPrev && prevPrev.populatedCellCount > 0) {
        score -= 2;
      }
    }

    if (score >= 6) return 'high';
    if (score >= 3) return 'medium';
    if (score >= 1) return 'low';
    return null;
  }

  private buildHeaderEvidence(
    summaries: RowSummary[],
    index: number,
    confidence: HeaderConfidence,
  ): string {
    const row = summaries[index];
    const parts: string[] = [];

    const textRatio = (row.textCellCount / row.populatedCellCount) * 100;
    parts.push(`${Math.round(textRatio)}% text cells`);

    if (row.hasMergedCells) parts.push('contains merged cells');

    const prev = summaries[index - 1];
    if (!prev || prev.populatedCellCount === 0) {
      parts.push('preceded by empty row or is first row');
    }

    const next = summaries[index + 1];
    if (next && next.textCellCount < next.populatedCellCount) {
      parts.push('followed by rows with numeric/formula data');
    }

    return `${confidence.toUpperCase()} confidence: ${parts.join('; ')}`;
  }

  private scoreHeaderCandidatesForSection(
    rowSummaries: RowSummary[],
    sectionStart: number,
    sectionEnd: number,
  ): HeaderCandidate[] {
    const candidates: HeaderCandidate[] = [];

    // Check the first 1-2 rows of the section
    for (let rowNum = sectionStart; rowNum <= Math.min(sectionStart + 1, sectionEnd); rowNum++) {
      const idx = rowSummaries.findIndex((r) => r.rowNumber === rowNum);
      if (idx === -1) continue;
      const confidence = this.scoreHeaderRow(rowSummaries, idx);
      if (confidence) {
        const row = rowSummaries[idx];
        candidates.push({
          rowNumber: rowNum,
          coordinates: row.coordinates,
          labels: row.values,
          populatedCellCount: row.populatedCellCount,
          textCellCount: row.textCellCount,
          hasMergedCells: row.hasMergedCells,
          confidence,
          evidence: this.buildHeaderEvidence(rowSummaries, idx, confidence),
        });
      }
    }

    return candidates;
  }

  // ─── Private: cell inspection (shared with ExcelInspectionService pattern) ─

  /**
   * Build an InspectedCell from an ExcelJS cell.
   *
   * Duplicates the same pattern as ExcelInspectionService.buildInspectedCell()
   * to keep services independent. If this logic grows complex, extract to a
   * shared utility in a later refactor.
   */
  private buildInspectedCell(
    sheetName: string,
    cell: ExcelJS.Cell,
  ): InspectedCell {
    const { displayValue, formula, type } = this.getCellDisplayValue(cell);
    const numFmt = cell.numFmt && cell.numFmt !== 'General' ? cell.numFmt : null;
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

  private getCellDisplayValue(cell: ExcelJS.Cell): {
    displayValue: string | number | boolean | null;
    formula: string | null;
    type: InspectedCellType;
  } {
    const vt = cell.type;

    if (vt === 0) return { displayValue: null, formula: null, type: 'empty' };

    if (vt === 1) {
      const rawVal = cell.value;
      const displayValue =
        rawVal !== null && rawVal !== undefined
          ? this.extractRawValue(rawVal)
          : null;
      return { displayValue, formula: null, type: 'empty' };
    }

    if (vt === 6) {
      const formulaValue = cell.value as ExcelJS.CellFormulaValue;
      const formulaStr =
        typeof formulaValue?.formula === 'string' ? formulaValue.formula : null;
      const result = formulaValue?.result;
      const displayValue =
        result !== undefined && result !== null
          ? this.extractRawValue(result)
          : null;
      return { displayValue, formula: formulaStr, type: 'formula' };
    }

    if (vt === 2) {
      return {
        displayValue: typeof cell.value === 'number' ? cell.value : null,
        formula: null,
        type: 'number',
      };
    }

    if (vt === 4) {
      const d = cell.value;
      return {
        displayValue: d instanceof Date ? d.toISOString() : String(d),
        formula: null,
        type: 'date',
      };
    }

    if (vt === 9) {
      return {
        displayValue: typeof cell.value === 'boolean' ? cell.value : null,
        formula: null,
        type: 'boolean',
      };
    }

    const strVal = this.extractStringValue(cell.value);
    return {
      displayValue: strVal,
      formula: null,
      type: strVal !== null ? 'string' : 'empty',
    };
  }

  private extractRawValue(
    val: ExcelJS.CellValue,
  ): string | number | boolean | null {
    if (val === null || val === undefined) return null;
    if (typeof val === 'string') return val;
    if (typeof val === 'number') return val;
    if (typeof val === 'boolean') return val;
    if (val instanceof Date) return val.toISOString();
    if (typeof val === 'object' && 'richText' in val) {
      return (val as ExcelJS.CellRichTextValue).richText
        .map((rt) => rt.text)
        .join('');
    }
    if (typeof val === 'object' && 'text' in val && 'hyperlink' in val) {
      return String((val as ExcelJS.CellHyperlinkValue).text);
    }
    if (typeof val === 'object' && 'formula' in val) {
      const r = (val as ExcelJS.CellFormulaValue).result;
      return r !== undefined && r !== null ? this.extractRawValue(r) : null;
    }
    return String(val);
  }

  private extractStringValue(val: ExcelJS.CellValue): string | null {
    const raw = this.extractRawValue(val);
    return raw === null || raw === undefined ? null : String(raw);
  }

  private buildMergeInfo(cell: ExcelJS.Cell) {
    if (!cell.isMerged) {
      return { status: 'none' as const, mergedRange: null, masterAddress: null };
    }
    const masterAddr = cell.master?.address ?? null;
    const isMaster = masterAddr === cell.address;
    const mergedRange = this.resolveMergeRange(cell, masterAddr);
    const status: 'master' | 'member' = isMaster ? 'master' : 'member';
    return { status, mergedRange, masterAddress: masterAddr };
  }

  private resolveMergeRange(
    cell: ExcelJS.Cell,
    masterAddr: string | null,
  ): string | null {
    try {
      const ws = (cell as any)._row?._worksheet;
      if (!ws || !ws._merges) return null;
      const effectiveMaster = masterAddr ?? cell.address;
      const mergeModel = ws._merges[effectiveMaster];
      if (!mergeModel?.model) return null;
      const { top, left, bottom, right } = mergeModel.model;
      return `${colNumberToLetter(left)}${top}:${colNumberToLetter(right)}${bottom}`;
    } catch {
      return null;
    }
  }

  // ─── Private: validation ──────────────────────────────────────────────────

  private validateSheetNameInput(sheetName: string): void {
    if (!sheetName || typeof sheetName !== 'string' || !sheetName.trim()) {
      throw new BadRequestException('Sheet name must be a non-empty string');
    }
  }
}

// ─── Internal helper type ─────────────────────────────────────────────────────

interface RowSummary {
  rowNumber: number;
  populatedCellCount: number;
  textCellCount: number;
  hasFormulas: boolean;
  hasMergedCells: boolean;
  coordinates: string[];
  values: (string | number | boolean | null)[];
  minCol: number;
  maxCol: number;
  isHidden: boolean;
}
