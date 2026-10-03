/**
 * coordinate.utils.ts
 *
 * Excel A1-style coordinate utilities.
 *
 * Supports:
 *   - Column letter ↔ number conversion (A=1, Z=26, AA=27, …)
 *   - A1 cell address parsing and building
 *   - Range string parsing (A1:C10)
 *   - Optional absolute references ($A$1, $A1, A$1) — $ is stripped
 *   - Strict validation — malformed inputs throw, never silent repair
 *
 * Row and column indices are always 1-based throughout this module.
 *
 * LIMITS:
 *   MAX_COL = 16384  (Excel XFD — column 16384)
 *   MAX_ROW = 1048576 (Excel row limit)
 *
 * These match the xlsx specification.
 */

// ─── Constants ───────────────────────────────────────────────────────────────

export const MAX_COL = 16_384;   // XFD in Excel
export const MAX_ROW = 1_048_576; // Maximum rows in xlsx

// ─── Types ───────────────────────────────────────────────────────────────────

export interface ParsedCell {
  /** 1-based row index */
  row: number;
  /** 1-based column index */
  col: number;
  /** Canonical A1 address string (no $ signs) */
  address: string;
}

export interface ParsedRange {
  /** Top-left cell */
  start: ParsedCell;
  /** Bottom-right cell */
  end: ParsedCell;
  /** Canonical range string e.g. "A1:C10" */
  address: string;
  /** Number of rows in the range (inclusive) */
  rows: number;
  /** Number of columns in the range (inclusive) */
  cols: number;
}

// ─── Column conversion ───────────────────────────────────────────────────────

/**
 * Convert a column letter string to a 1-based column index.
 *
 * Examples:
 *   "A"  → 1
 *   "Z"  → 26
 *   "AA" → 27
 *   "AZ" → 52
 *   "BA" → 53
 *   "XFD" → 16384
 *
 * @throws {Error} if the input is not a valid column string
 */
export function colLetterToNumber(col: string): number {
  if (!col || typeof col !== 'string') {
    throw new Error(`Invalid column: ${JSON.stringify(col)}`);
  }

  // Strip leading/trailing whitespace
  const cleaned = col.trim().toUpperCase();

  if (!/^[A-Z]+$/.test(cleaned)) {
    throw new Error(
      `Invalid column letters "${col}": must contain only A-Z characters`,
    );
  }

  let result = 0;
  for (let i = 0; i < cleaned.length; i++) {
    result = result * 26 + (cleaned.charCodeAt(i) - 64); // 'A' = 65, so -64 gives 1
  }

  if (result > MAX_COL) {
    throw new Error(
      `Column "${col}" (${result}) exceeds maximum allowed column ${MAX_COL} (XFD)`,
    );
  }

  return result;
}

/**
 * Convert a 1-based column index to a column letter string.
 *
 * Examples:
 *   1  → "A"
 *   26 → "Z"
 *   27 → "AA"
 *   52 → "AZ"
 *   53 → "BA"
 *
 * @throws {Error} if the index is out of range
 */
export function colNumberToLetter(col: number): string {
  if (!Number.isInteger(col) || col < 1 || col > MAX_COL) {
    throw new Error(
      `Invalid column index ${col}: must be an integer between 1 and ${MAX_COL}`,
    );
  }

  let result = '';
  let n = col;

  while (n > 0) {
    const rem = (n - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    n = Math.floor((n - 1) / 26);
  }

  return result;
}

// ─── Cell address parsing ────────────────────────────────────────────────────

/**
 * Parse an Excel A1-style cell address into row and column numbers.
 *
 * Accepts optional absolute reference markers ($):
 *   "A1", "$A1", "A$1", "$A$1" — all accepted, $ is stripped.
 *
 * Examples:
 *   "A1"   → { row: 1, col: 1, address: "A1" }
 *   "G42"  → { row: 42, col: 7, address: "G42" }
 *   "AA10" → { row: 10, col: 27, address: "AA10" }
 *   "$A$1" → { row: 1, col: 1, address: "A1" }
 *
 * @throws {Error} for any malformed input
 */
export function parseCell(address: string): ParsedCell {
  if (!address || typeof address !== 'string') {
    throw new Error(`Invalid cell address: ${JSON.stringify(address)}`);
  }

  // Strip $ signs for absolute references
  const cleaned = address.trim().replace(/\$/g, '').toUpperCase();

  if (!cleaned) {
    throw new Error(`Cell address is empty after stripping: "${address}"`);
  }

  // Match: one or more letters followed by one or more digits
  const match = cleaned.match(/^([A-Z]+)(\d+)$/);

  if (!match) {
    throw new Error(
      `Invalid cell address "${address}": must be in A1 format (letters then digits, e.g. "G42")`,
    );
  }

  const [, colStr, rowStr] = match;

  const row = parseInt(rowStr, 10);
  if (row < 1 || row > MAX_ROW) {
    throw new Error(
      `Cell address "${address}": row ${row} is out of range (1–${MAX_ROW})`,
    );
  }

  const col = colLetterToNumber(colStr); // throws if out of range

  return {
    row,
    col,
    address: `${colStr}${row}`,
  };
}

/**
 * Build an A1 address string from row and column numbers.
 *
 * Examples:
 *   (1, 1) → "A1"
 *   (42, 7) → "G42"
 *
 * @throws {Error} for invalid row/col
 */
export function buildCellAddress(row: number, col: number): string {
  if (!Number.isInteger(row) || row < 1 || row > MAX_ROW) {
    throw new Error(
      `Invalid row ${row}: must be an integer between 1 and ${MAX_ROW}`,
    );
  }
  return `${colNumberToLetter(col)}${row}`;
}

// ─── Range parsing ────────────────────────────────────────────────────────────

/**
 * Parse a range string into start and end cell coordinates.
 *
 * Accepts:
 *   "A1:C10"
 *   "$A$1:$C$10"
 *   "G42:J50"
 *   "AA5:AD20"
 *
 * Rules:
 *   - start must be top-left of the range (row ≤ end row, col ≤ end col)
 *   - If the range is not top-left-to-bottom-right, it is normalized
 *
 * @throws {Error} for any malformed range string
 */
export function parseRange(range: string): ParsedRange {
  if (!range || typeof range !== 'string') {
    throw new Error(`Invalid range: ${JSON.stringify(range)}`);
  }

  const cleaned = range.trim();

  const colonIdx = cleaned.indexOf(':');
  if (colonIdx === -1) {
    throw new Error(
      `Invalid range "${range}": must contain a colon separating start and end cells (e.g. "A1:C10")`,
    );
  }

  const startStr = cleaned.slice(0, colonIdx);
  const endStr = cleaned.slice(colonIdx + 1);

  if (!startStr || !endStr) {
    throw new Error(
      `Invalid range "${range}": both start and end cell addresses are required`,
    );
  }

  const start = parseCell(startStr);
  const end = parseCell(endStr);

  // Normalize to ensure start ≤ end
  const topRow = Math.min(start.row, end.row);
  const leftCol = Math.min(start.col, end.col);
  const bottomRow = Math.max(start.row, end.row);
  const rightCol = Math.max(start.col, end.col);

  const normalizedStart: ParsedCell = {
    row: topRow,
    col: leftCol,
    address: buildCellAddress(topRow, leftCol),
  };
  const normalizedEnd: ParsedCell = {
    row: bottomRow,
    col: rightCol,
    address: buildCellAddress(bottomRow, rightCol),
  };

  return {
    start: normalizedStart,
    end: normalizedEnd,
    address: `${normalizedStart.address}:${normalizedEnd.address}`,
    rows: bottomRow - topRow + 1,
    cols: rightCol - leftCol + 1,
  };
}

/**
 * Validate a cell address string without throwing.
 * Returns true if valid, false otherwise.
 */
export function isCellAddressValid(address: string): boolean {
  try {
    parseCell(address);
    return true;
  } catch {
    return false;
  }
}

/**
 * Validate a range string without throwing.
 * Returns true if valid, false otherwise.
 */
export function isRangeValid(range: string): boolean {
  try {
    parseRange(range);
    return true;
  } catch {
    return false;
  }
}
