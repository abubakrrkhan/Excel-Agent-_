/**
 * structure.types.ts
 *
 * TypeScript interfaces for Phase 1 Step 2 — Workbook Intelligence &
 * Structural Inspection.
 *
 * These are READ-ONLY structural analysis results.
 * They are evidence artifacts for the future anti-hallucination agent.
 *
 * No write operations are represented here.
 * All interfaces are additive — inspection.types.ts is not replaced.
 */

import { InspectedCell } from './inspection.types';

// ─── Sheet structure ──────────────────────────────────────────────────────────

/**
 * A summary of a single merged cell range at the sheet level.
 * Uses the public ws.model?.merges[] API.
 */
export interface SheetMergeRange {
  /** Canonical range address e.g. "A1:C1" */
  range: string;
  /** A1 address of the master (top-left) cell */
  masterAddress: string;
  /** Number of rows this merge spans */
  rowSpan: number;
  /** Number of columns this merge spans */
  colSpan: number;
}

/**
 * Top-level sheet-structure summary.
 * Used to give the future agent a rapid overview before detailed inspection.
 */
export interface SheetStructure {
  sheetName: string;

  /** Used bounds — where meaningful content starts and ends */
  usedBounds: {
    /** First row with any content (1-based), or null if sheet is empty */
    firstRow: number | null;
    /** Last row with any content (1-based), or null if sheet is empty */
    lastRow: number | null;
    /** First column with any content (1-based), or null if sheet is empty */
    firstCol: number | null;
    /** Last column with any content (1-based), or null if sheet is empty */
    lastCol: number | null;
    /** A1:XN style address of used region, or null if empty */
    address: string | null;
  };

  /** Counts of meaningful row types */
  rowCount: {
    /** Total rows ExcelJS reports (includes trailing empty rows) */
    total: number;
    /** Rows that actually contain at least one non-empty cell */
    populated: number;
    /** Rows explicitly marked hidden */
    hidden: number;
  };

  /** Counts of meaningful column types */
  colCount: {
    /** Total columns ExcelJS reports */
    total: number;
    /** Columns that actually contain at least one non-empty cell */
    populated: number;
    /** Columns explicitly marked hidden */
    hidden: number;
  };

  /** Counts of cell types across the entire sheet */
  cellCounts: {
    nonEmpty: number;
    text: number;
    numeric: number;
    date: number;
    formula: number;
    boolean: number;
  };

  /** All merged ranges on this sheet (from public ws.model?.merges) */
  mergedRanges: SheetMergeRange[];

  /** 1-based row numbers that are explicitly hidden */
  hiddenRows: number[];

  /** 1-based column numbers that are explicitly hidden */
  hiddenCols: number[];
}

// ─── Header detection ─────────────────────────────────────────────────────────

/**
 * Confidence level for a header candidate.
 *
 * 'high'   — strong evidence (many text cells, follows an empty row, merged spans)
 * 'medium' — moderate evidence
 * 'low'    — weak evidence (could be a data row)
 */
export type HeaderConfidence = 'high' | 'medium' | 'low';

/**
 * A candidate header row identified by structural analysis.
 * Multiple candidates may exist in one sheet.
 */
export interface HeaderCandidate {
  /** 1-based row number */
  rowNumber: number;
  /** A1 addresses of non-empty cells in this row */
  coordinates: string[];
  /** String values found in this row (for display/agent context) */
  labels: (string | number | boolean | null)[];
  /** How many non-empty cells this row contains */
  populatedCellCount: number;
  /** How many of those cells are text (vs numeric/formula) */
  textCellCount: number;
  /** True if any cells in this row are merged */
  hasMergedCells: boolean;
  /** Confidence assessment */
  confidence: HeaderConfidence;
  /**
   * Human-readable evidence summary.
   * Explains WHY this row was identified as a header candidate.
   * The agent can use this to reason about which candidate is correct.
   */
  evidence: string;
}

/**
 * Result of header detection for a worksheet.
 */
export interface HeaderDetectionResult {
  sheetName: string;
  /**
   * All header candidates, ordered by confidence descending, then by row number.
   * May be empty if no candidates were found.
   */
  candidates: HeaderCandidate[];
  /**
   * True if more than one candidate was found.
   * The future agent must not assume the first candidate is correct
   * if this is true — it must inspect context and ask for clarification.
   */
  hasMultipleCandidates: boolean;
}

// ─── Table / section detection ────────────────────────────────────────────────

/**
 * A candidate table or section region within a worksheet.
 *
 * A region is a contiguous block of non-empty cells,
 * separated from other regions by at least one fully empty row
 * or a significant gap.
 */
export interface SectionCandidate {
  /** Sequential index of this candidate (0-based) within the sheet */
  index: number;
  /** A1 address of the top-left cell of this region */
  startCoordinate: string;
  /** A1 address of the bottom-right cell of this region */
  endCoordinate: string;
  /** Canonical range address e.g. "A4:D20" */
  rangeAddress: string;
  /** Number of rows in this region */
  rows: number;
  /** Number of columns in this region */
  cols: number;
  /** Approximate number of non-empty cells */
  nonEmptyCells: number;
  /** Whether any formula cells exist in this region */
  hasFormulas: boolean;
  /** Whether any merged cells exist in this region */
  hasMergedCells: boolean;
  /**
   * Header candidates within or immediately above this region.
   * May be empty if no header was detected.
   */
  headerCandidates: HeaderCandidate[];
}

/**
 * Result of section/table detection for a worksheet.
 */
export interface SectionDetectionResult {
  sheetName: string;
  /** All candidate sections, ordered by start row */
  sections: SectionCandidate[];
}

// ─── Cell context ─────────────────────────────────────────────────────────────

/**
 * Options for surrounding context inspection.
 */
export interface CellContextOptions {
  /**
   * Number of rows above and below the target to include.
   * Default: 3
   */
  rowRadius?: number;
  /**
   * Number of columns left and right of the target to include.
   * Default: 3
   */
  colRadius?: number;
}

/**
 * The full context window around an inspected target cell.
 */
export interface CellContext {
  sheetName: string;
  /** The target cell that was inspected */
  targetCell: InspectedCell;
  /** Effective context window used (after clamping to sheet bounds) */
  contextRange: {
    startRow: number;
    endRow: number;
    startCol: number;
    endCol: number;
    address: string;
  };
  /**
   * All cells in the context window, ordered row-by-row, left-to-right.
   * Includes the target cell.
   */
  cells: InspectedCell[];
  /**
   * Non-empty label cells found in the same row as the target.
   * Useful for identifying what the target cell "belongs to".
   */
  rowLabels: InspectedCell[];
  /**
   * Non-empty cells found in the same column as the target.
   * Useful for understanding the column's purpose.
   */
  colLabels: InspectedCell[];
}

// ─── Enhanced search match ────────────────────────────────────────────────────

/**
 * Optional surrounding evidence for a search match.
 * Added to SearchMatch when includeContext=true is requested.
 */
export interface SearchMatchContext {
  /** The inspected cell with full detail (value, formula, type, merge) */
  cell: InspectedCell;
  /**
   * Labels found in the same row as the match.
   * Helps the agent understand what section/field the match belongs to.
   */
  rowLabels: Array<{
    coordinate: string;
    value: string | number | boolean | null;
    type: string;
  }>;
  /**
   * Labels found in the rows immediately above the match (up to 2 rows).
   * Often contains section headings.
   */
  nearbyAbove: Array<{
    coordinate: string;
    value: string | number | boolean | null;
    rowNumber: number;
  }>;
}

// ─── Limits and defaults ──────────────────────────────────────────────────────

export const DEFAULT_CELL_CONTEXT_OPTIONS: Required<CellContextOptions> = {
  rowRadius: 3,
  colRadius: 3,
};

/**
 * Maximum row radius allowed for context inspection.
 * Prevents accidental huge reads.
 */
export const MAX_CELL_CONTEXT_ROW_RADIUS = 20;
export const MAX_CELL_CONTEXT_COL_RADIUS = 20;

/**
 * Maximum gap (consecutive empty rows) before treating as a section boundary.
 */
export const SECTION_SPLIT_GAP = 2;

/**
 * Minimum number of non-empty cells for a region to qualify as a section candidate.
 */
export const SECTION_MIN_CELLS = 2;
