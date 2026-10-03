/**
 * inspection.types.ts
 *
 * TypeScript interfaces for Phase 1 Deep Workbook Inspection.
 *
 * These types are READ-ONLY inspection results.
 * They are the future evidence API for the anti-hallucination agent.
 *
 * No write operations are represented here.
 */

// ─── Cell type ───────────────────────────────────────────────────────────────

/**
 * Normalized cell value type derived from ExcelJS ValueType enum.
 *
 * Mapping:
 *   ExcelJS.ValueType.Null    (0) → 'empty'
 *   ExcelJS.ValueType.Merge   (1) → handled as merged-cell membership
 *   ExcelJS.ValueType.Number  (2) → 'number'
 *   ExcelJS.ValueType.String  (3) → 'string'
 *   ExcelJS.ValueType.Date    (4) → 'date'
 *   ExcelJS.ValueType.Formula (6) → 'formula'
 *   ExcelJS.ValueType.Boolean (9) → 'boolean'
 *   All others                    → 'string' (safe fallback)
 */
export type InspectedCellType =
  | 'string'
  | 'number'
  | 'boolean'
  | 'date'
  | 'formula'
  | 'empty';

// ─── Merge info ──────────────────────────────────────────────────────────────

/**
 * Describes a cell's relationship to a merged cell range.
 *
 * 'none'   — cell is not part of any merge
 * 'master' — cell is the top-left cell of a merged range (holds the value)
 * 'member' — cell is inside a merged range but is not the master
 */
export type MergeStatus = 'none' | 'master' | 'member';

export interface MergeInfo {
  /** Whether and how this cell participates in a merged range */
  status: MergeStatus;
  /**
   * The full merged range in A1:B2 format.
   * Present when status is 'master' or 'member'.
   */
  mergedRange: string | null;
  /**
   * The A1 address of the top-left (master) cell of the merge.
   * Present when status is 'master' or 'member'.
   * For 'master' cells, equals the cell's own address.
   */
  masterAddress: string | null;
}

// ─── Inspected cell ──────────────────────────────────────────────────────────

export interface InspectedCell {
  /** Sheet name */
  sheetName: string;
  /** A1 address of this cell (e.g. "G42") */
  coordinate: string;
  /**
   * Resolved display value.
   * For formula cells, this is the cached result (not the formula string).
   * For empty/merged-member cells, this is null.
   */
  value: string | number | boolean | null;
  /**
   * Raw formula string if the cell contains a formula.
   * e.g. "SUM(A1:A10)"
   * null if the cell is not a formula cell.
   */
  formula: string | null;
  /** Normalized cell type */
  type: InspectedCellType;
  /**
   * Excel number format string if set (e.g. "0.00%", "dd/mm/yyyy").
   * null if no custom format is applied.
   */
  numFmt: string | null;
  /** Merge information */
  merge: MergeInfo;
}

// ─── Range reader ─────────────────────────────────────────────────────────────

export interface RangeReaderOptions {
  /**
   * Maximum number of cells allowed in the requested range.
   * Requests exceeding this are rejected.
   * Default: 10000
   */
  maxCells?: number;
}

export interface InspectedRange {
  sheetName: string;
  /** Canonical range address (e.g. "G40:J45") */
  range: string;
  dimensions: {
    rows: number;
    cols: number;
    totalCells: number;
  };
  cells: InspectedCell[];
}

// ─── Workbook search ──────────────────────────────────────────────────────────

export interface SearchOptions {
  /**
   * Whether to match case-insensitively.
   * Default: true (case-insensitive)
   */
  caseInsensitive?: boolean;
  /**
   * Restrict the search to a specific worksheet by name.
   * When absent, all worksheets are searched.
   */
  sheetName?: string;
  /**
   * Maximum number of matches to return.
   * Default: 50 — prevents unbounded responses.
   */
  maxMatches?: number;
  /**
   * Maximum number of cells to scan across the entire workbook.
   * Default: 100000 — protects against huge workbooks.
   */
  maxCellsScanned?: number;
  /**
   * When true, each match includes surrounding row labels and nearby headings.
   * This enriches the evidence available to the agent for disambiguation.
   * Default: false — context enrichment is opt-in to avoid overhead on simple searches.
   */
  includeContext?: boolean;
}

export interface SearchMatch {
  /** Sheet where the match was found */
  sheetName: string;
  /** A1 address of the matched cell */
  coordinate: string;
  /**
   * The display value of the matched cell as a string.
   * For formula cells, this is the cached result.
   */
  value: string;
  /** The normalized cell type */
  type: InspectedCellType;
  /**
   * Surrounding evidence for this match.
   * Only populated when SearchOptions.includeContext = true.
   * Contains row labels, nearby headings, and full cell detail.
   * The agent uses this to disambiguate between multiple matches of the same label.
   */
  context?: import('./structure.types').SearchMatchContext;
}

export interface SearchResult {
  /** Original query string */
  query: string;
  /** Effective options used (with defaults applied) */
  options: Required<SearchOptions>;
  /** All matches found within the limits */
  matches: SearchMatch[];
  /** Number of cells scanned before stopping */
  cellsScanned: number;
  /**
   * True if the scan hit maxCellsScanned or maxMatches before finishing
   * all worksheets. The caller should know results may be incomplete.
   */
  truncated: boolean;
}

// ─── Search/range limits (defaults) ─────────────────────────────────────────

export const DEFAULT_SEARCH_OPTS: Required<SearchOptions> = {
  caseInsensitive: true,
  sheetName: '',          // empty string = all sheets
  maxMatches: 50,
  maxCellsScanned: 100_000,
  includeContext: false,  // opt-in — context adds overhead
};

export const DEFAULT_MAX_RANGE_CELLS = 10_000;
