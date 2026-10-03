/**
 * operation.types.ts
 *
 * Strongly-typed parameter contracts for all supported operations.
 *
 * Phase 2 Step 1 additions:
 *   - INSPECT_WORKBOOK_STRUCTURE
 *   - DETECT_HEADERS
 *   - DETECT_SECTIONS
 *   - GET_CELL_CONTEXT
 *   - SET_CELL_BY_LABEL
 *   - allowFormulaOverwrite on SET_CELL
 *
 * Existing operation param shapes preserved exactly.
 */

// ─── Write operation params ───────────────────────────────────────────────────

export interface SetCellParams {
  /** 1-based row index */
  row: number;
  /** 1-based column index */
  col: number;
  /** Value to set */
  value: string | number | boolean | null;
  /**
   * When true, explicitly authorizes overwriting an existing formula cell.
   * Default: false — formula cells are protected by default.
   *
   * Without this flag, attempting to overwrite a formula cell raises a
   * FormulaOverwriteError. The caller must pass `allowFormulaOverwrite: true`
   * to confirm intentional formula replacement.
   */
  allowFormulaOverwrite?: boolean;
}

export interface SetCellRangeParams {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
  value: string | number | boolean | null;
  /**
   * When true, explicitly authorizes overwriting formula cells within the range.
   * Default: false — formula cells are protected by default.
   *
   * The entire range is pre-scanned for formula cells before any write occurs.
   * If protected formula cells are found and this flag is false, the operation
   * throws without modifying any cell (atomic — no partial writes).
   */
  allowFormulaOverwrite?: boolean;
}

export interface AddRowParams {
  values: (string | number | boolean | null)[];
}

export interface DeleteRowParams {
  row: number;
}

export interface AddFormulaParams {
  row: number;
  col: number;
  formula: string;
}

export interface FormatCellParams {
  row: number;
  col: number;
  style: Record<string, unknown>;
}

export interface AddColumnHeaderParams {
  col: number;
  header: string;
}

export interface SortSheetParams {
  col: number;
  order: 'asc' | 'desc';
}

// ─── Inspection operation params ──────────────────────────────────────────────

export interface InspectWorkbookStructureParams {
  /** Name of the worksheet to inspect */
  sheetName: string;
}

export interface DetectHeadersParams {
  /** Name of the worksheet to analyse */
  sheetName: string;
}

export interface DetectSectionsParams {
  /** Name of the worksheet to analyse */
  sheetName: string;
}

export interface GetCellContextParams {
  /** Name of the worksheet */
  sheetName: string;
  /** A1-style coordinate of the target cell (e.g. "G42") */
  coordinate: string;
  /** Number of rows above/below target to include (default: 3, max: 20) */
  rowRadius?: number;
  /** Number of columns left/right of target to include (default: 3, max: 20) */
  colRadius?: number;
}

// ─── SET_CELL_BY_LABEL params and result ─────────────────────────────────────

export interface SetCellByLabelParams {
  /**
   * The label text to search for in the workbook.
   * Example: "Management Fee"
   */
  label: string;
  /**
   * The value to write to the resolved cell.
   * Null clears the cell.
   */
  value: string | number | boolean | null;
  /**
   * Optionally restrict the search to one worksheet.
   * When absent, all worksheets are searched.
   */
  sheetName?: string;
  /**
   * Column offset from the label cell to the value cell.
   * Default: 1 (the cell immediately to the right of the label).
   *
   * Example: label="Management Fee" in A5, valueColOffset=1 → writes B5.
   */
  valueColOffset?: number;
  /**
   * Authorizes overwriting a formula cell at the resolved target.
   * Default: false.
   */
  allowFormulaOverwrite?: boolean;
}

/**
 * Resolution status for SET_CELL_BY_LABEL.
 *
 * 'resolved'            — exactly one match found, evidence sufficient
 * 'ambiguous'           — multiple matches found; cannot safely choose
 * 'not_found'           — label not found in workbook
 * 'insufficient_evidence' — match found but context is ambiguous; caller must inspect further
 */
export type LabelResolutionStatus =
  | 'resolved'
  | 'ambiguous'
  | 'not_found'
  | 'insufficient_evidence';

export interface LabelMatch {
  sheetName: string;
  labelCoordinate: string;
  /** Coordinate of the target value cell (labelCoordinate + offset) */
  valueCoordinate: string;
  labelValue: string;
}

export interface SetCellByLabelResult {
  status: LabelResolutionStatus;
  /** Human-readable explanation of why this status was returned */
  message: string;
  /** All candidate matches found (always populated, even for 'ambiguous') */
  candidates: LabelMatch[];
  /**
   * The resolved target — only populated when status = 'resolved'.
   * The caller must use this coordinate for the actual SET_CELL write.
   */
  resolvedTarget: LabelMatch | null;
}

// ─── Formula overwrite error ──────────────────────────────────────────────────

export interface FormulaOverwriteError {
  /** The coordinate that was blocked */
  coordinate: string;
  /** The formula currently in the cell */
  existingFormula: string;
  /** The cached result of the formula */
  cachedResult: string | number | null;
  /** Instructions for the caller */
  message: string;
}

// ─── Operation type groups ────────────────────────────────────────────────────

/**
 * Operations that modify the workbook.
 * These require a workbook buffer and produce a modified workbook.
 */
export const WRITE_OPERATIONS = [
  'SET_CELL',
  'SET_CELL_RANGE',
  'ADD_ROW',
  'DELETE_ROW',
  'ADD_FORMULA',
  'FORMAT_CELL',
  'ADD_COLUMN_HEADER',
  'SORT_SHEET',
] as const;

/**
 * Operations that only read the workbook.
 * These require a workbook buffer but never modify it.
 */
export const READ_OPERATIONS = [
  'INSPECT_WORKBOOK_STRUCTURE',
  'DETECT_HEADERS',
  'DETECT_SECTIONS',
  'GET_CELL_CONTEXT',
] as const;

/**
 * Hybrid operations that read first, then optionally write.
 * These require a workbook buffer and may produce a modified workbook.
 */
export const HYBRID_OPERATIONS = [
  'SET_CELL_BY_LABEL',
] as const;

export type WriteOperationType = typeof WRITE_OPERATIONS[number];
export type ReadOperationType = typeof READ_OPERATIONS[number];
export type HybridOperationType = typeof HYBRID_OPERATIONS[number];
export type AnyOperationType = WriteOperationType | ReadOperationType | HybridOperationType;

/** All supported operation types as a flat array */
export const ALL_OPERATION_TYPES: string[] = [
  ...WRITE_OPERATIONS,
  ...READ_OPERATIONS,
  ...HYBRID_OPERATIONS,
];
