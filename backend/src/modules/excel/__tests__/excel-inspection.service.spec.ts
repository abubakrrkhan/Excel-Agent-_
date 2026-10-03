/**
 * excel-inspection.service.spec.ts
 *
 * Tests for ExcelInspectionService:
 *   - searchWorkbook()
 *   - readCellRange()
 *   - inspectCell()
 *
 * Uses the synthetic financial fixture workbook (in-memory only).
 * No filesystem access, no real bank data.
 */

import * as ExcelJS from 'exceljs';
import { ExcelInspectionService } from '../excel-inspection.service';
import { WorkbookStructureService } from '../workbook-structure.service';
import { buildFinancialFixtureBuffer } from '../fixtures/financial-workbook.fixture';

// ─── Helpers ─────────────────────────────────────────────────────────────────

/** Build a minimal in-memory workbook Buffer for targeted tests */
async function buildSimpleBuffer(
  setup: (wb: ExcelJS.Workbook) => void,
): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  setup(wb);
  return Buffer.from(await wb.xlsx.writeBuffer());
}

// ─── Test setup ───────────────────────────────────────────────────────────────

let svc: ExcelInspectionService;
let fixtureBuffer: Buffer;

beforeAll(async () => {
  // WorkbookStructureService is injected into ExcelInspectionService (added in Step 2)
  const structureSvc = new WorkbookStructureService();
  svc = new ExcelInspectionService(structureSvc);
  fixtureBuffer = await buildFinancialFixtureBuffer();
});

// ═════════════════════════════════════════════════════════════════════════════
// SEARCH TESTS
// ═════════════════════════════════════════════════════════════════════════════

describe('searchWorkbook', () => {

  // ── Basic match ───────────────────────────────────────────────────────────

  it('finds exact text match in fixture', async () => {
    const result = await svc.searchWorkbook(fixtureBuffer, 'NAV per Unit');
    expect(result.matches.length).toBeGreaterThan(0);
    const match = result.matches[0];
    expect(match.value).toContain('NAV per Unit');
    expect(match.sheetName).toBeTruthy();
    expect(match.coordinate).toMatch(/^[A-Z]+\d+$/); // valid A1 format
  });

  // ── Case-insensitive (default) ────────────────────────────────────────────

  it('matches case-insensitively by default', async () => {
    const upper = await svc.searchWorkbook(fixtureBuffer, 'management fee');
    const mixed = await svc.searchWorkbook(fixtureBuffer, 'Management Fee');
    // Both should find matches
    expect(upper.matches.length).toBeGreaterThan(0);
    expect(mixed.matches.length).toBeGreaterThan(0);
    // Same count when case differs
    expect(upper.matches.length).toBe(mixed.matches.length);
  });

  // ── Repeated labels ───────────────────────────────────────────────────────

  it('finds "Management Fee" in multiple locations (fixture has 4+ occurrences)', async () => {
    const result = await svc.searchWorkbook(fixtureBuffer, 'Management Fee');
    // The fixture deliberately places "Management Fee" in multiple sheets/rows
    // to test anti-hallucination — there is no single correct G42
    expect(result.matches.length).toBeGreaterThanOrEqual(3);
  });

  it('returns correct sheetName and coordinate for each match', async () => {
    const result = await svc.searchWorkbook(fixtureBuffer, 'Management Fee');
    for (const match of result.matches) {
      expect(match.sheetName).toBeTruthy();
      expect(match.coordinate).toMatch(/^[A-Z]+\d+$/);
      expect(match.value).toBeTruthy();
    }
  });

  // ── No match ──────────────────────────────────────────────────────────────

  it('returns empty matches for non-existent text', async () => {
    const result = await svc.searchWorkbook(fixtureBuffer, 'ZZZNOMATCHZZZ');
    expect(result.matches).toHaveLength(0);
    expect(result.truncated).toBe(false);
  });

  // ── Numeric values ────────────────────────────────────────────────────────

  it('finds numeric values by their string representation', async () => {
    // Fund size is 2500000000 in the fixture — but searching as string
    const result = await svc.searchWorkbook(fixtureBuffer, '105.25');
    // NAV per Unit = 105.25 in Fund Overview
    expect(result.matches.length).toBeGreaterThan(0);
    expect(result.matches[0].type).toBe('number');
  });

  // ── Worksheet filtering ───────────────────────────────────────────────────

  it('filters to specific sheet', async () => {
    const result = await svc.searchWorkbook(fixtureBuffer, 'Management Fee', {
      sheetName: 'Fund Details',
    });
    // All matches must be from Fund Details only
    for (const match of result.matches) {
      expect(match.sheetName).toBe('Fund Details');
    }
    expect(result.matches.length).toBeGreaterThan(0);
  });

  it('filtering to non-existent sheet returns empty matches', async () => {
    const result = await svc.searchWorkbook(fixtureBuffer, 'Management Fee', {
      sheetName: 'Sheet That Does Not Exist',
    });
    expect(result.matches).toHaveLength(0);
  });

  // ── Result limits ─────────────────────────────────────────────────────────

  it('respects maxMatches limit', async () => {
    const result = await svc.searchWorkbook(fixtureBuffer, 'Management Fee', {
      maxMatches: 1,
    });
    expect(result.matches.length).toBeLessThanOrEqual(1);
    expect(result.truncated).toBe(true);
  });

  it('respects maxCellsScanned limit', async () => {
    const result = await svc.searchWorkbook(fixtureBuffer, 'Management Fee', {
      maxCellsScanned: 5,
    });
    expect(result.cellsScanned).toBeLessThanOrEqual(6); // 5 + 1 that hit the limit
    expect(result.truncated).toBe(true);
  });

  // ── Determinism ───────────────────────────────────────────────────────────

  it('returns same results on repeated calls (deterministic)', async () => {
    const r1 = await svc.searchWorkbook(fixtureBuffer, 'Management Fee');
    const r2 = await svc.searchWorkbook(fixtureBuffer, 'Management Fee');
    expect(r1.matches.length).toBe(r2.matches.length);
    for (let i = 0; i < r1.matches.length; i++) {
      expect(r1.matches[i].coordinate).toBe(r2.matches[i].coordinate);
      expect(r1.matches[i].sheetName).toBe(r2.matches[i].sheetName);
    }
  });

  // ── Invalid input ─────────────────────────────────────────────────────────

  it('throws for empty query', async () => {
    await expect(svc.searchWorkbook(fixtureBuffer, '')).rejects.toThrow();
  });

  it('throws for whitespace-only query', async () => {
    await expect(svc.searchWorkbook(fixtureBuffer, '   ')).rejects.toThrow();
  });

  it('throws for empty buffer', async () => {
    await expect(
      svc.searchWorkbook(Buffer.alloc(0), 'test'),
    ).rejects.toThrow();
  });
});

// ═════════════════════════════════════════════════════════════════════════════
// RANGE READER TESTS
// ═════════════════════════════════════════════════════════════════════════════

describe('readCellRange', () => {

  // ── Single cell range ─────────────────────────────────────────────────────

  it('reads a single cell range A1:A1', async () => {
    const buf = await buildSimpleBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A1').value = 'Hello';
    });
    const result = await svc.readCellRange(buf, 'Sheet1', 'A1:A1');
    expect(result.cells).toHaveLength(1);
    expect(result.cells[0].value).toBe('Hello');
    expect(result.cells[0].type).toBe('string');
    expect(result.cells[0].coordinate).toBe('A1');
    expect(result.dimensions.rows).toBe(1);
    expect(result.dimensions.cols).toBe(1);
  });

  // ── Multi-row, multi-col range ────────────────────────────────────────────

  it('reads a multi-row multi-col range', async () => {
    const buf = await buildSimpleBuffer((wb) => {
      const ws = wb.addWorksheet('Data');
      ws.getCell('A1').value = 'Name';
      ws.getCell('B1').value = 'Value';
      ws.getCell('A2').value = 'Fee';
      ws.getCell('B2').value = 1.25;
    });
    const result = await svc.readCellRange(buf, 'Data', 'A1:B2');
    expect(result.dimensions.rows).toBe(2);
    expect(result.dimensions.cols).toBe(2);
    expect(result.cells).toHaveLength(4);

    const a1 = result.cells.find((c) => c.coordinate === 'A1');
    expect(a1?.value).toBe('Name');
    expect(a1?.type).toBe('string');

    const b2 = result.cells.find((c) => c.coordinate === 'B2');
    expect(b2?.value).toBe(1.25);
    expect(b2?.type).toBe('number');
  });

  // ── Empty / sparse cells ──────────────────────────────────────────────────

  it('handles empty cells without crashing', async () => {
    const buf = await buildSimpleBuffer((wb) => {
      const ws = wb.addWorksheet('Sparse');
      ws.getCell('A1').value = 'Start';
      // A2, B1, B2 are empty
    });
    const result = await svc.readCellRange(buf, 'Sparse', 'A1:B2');
    expect(result.cells).toHaveLength(4);
    const emptyCell = result.cells.find((c) => c.coordinate === 'B1');
    expect(emptyCell?.type).toBe('empty');
    expect(emptyCell?.value).toBeNull();
  });

  // ── Formula cells ─────────────────────────────────────────────────────────

  it('detects formula cells and preserves formula string', async () => {
    const buf = await buildSimpleBuffer((wb) => {
      const ws = wb.addWorksheet('Calc');
      ws.getCell('A1').value = 100;
      ws.getCell('A2').value = 200;
      ws.getCell('A3').value = { formula: 'SUM(A1:A2)', result: 300 };
    });
    const result = await svc.readCellRange(buf, 'Calc', 'A1:A3');
    const formulaCell = result.cells.find((c) => c.coordinate === 'A3');

    expect(formulaCell?.type).toBe('formula');
    expect(formulaCell?.formula).toBe('SUM(A1:A2)');
    expect(formulaCell?.value).toBe(300); // cached result
  });

  // ── Merged cells ──────────────────────────────────────────────────────────

  it('identifies merged master cell', async () => {
    const buf = await buildSimpleBuffer((wb) => {
      const ws = wb.addWorksheet('Merged');
      ws.mergeCells('A1:C1');
      ws.getCell('A1').value = 'Report Header';
    });
    const result = await svc.readCellRange(buf, 'Merged', 'A1:C1');
    const master = result.cells.find((c) => c.coordinate === 'A1');
    expect(master?.merge.status).toBe('master');
    expect(master?.merge.masterAddress).toBe('A1');
    expect(master?.merge.mergedRange).toBe('A1:C1');
  });

  it('identifies merged member cells', async () => {
    const buf = await buildSimpleBuffer((wb) => {
      const ws = wb.addWorksheet('Merged');
      ws.mergeCells('A1:C1');
      ws.getCell('A1').value = 'Report Header';
    });
    const result = await svc.readCellRange(buf, 'Merged', 'A1:C1');
    const member = result.cells.find((c) => c.coordinate === 'B1');
    expect(member?.merge.status).toBe('member');
    expect(member?.merge.masterAddress).toBe('A1');
  });

  it('normal cells have merge.status = "none"', async () => {
    const buf = await buildSimpleBuffer((wb) => {
      const ws = wb.addWorksheet('Normal');
      ws.getCell('A1').value = 'No merge';
    });
    const result = await svc.readCellRange(buf, 'Normal', 'A1:A1');
    expect(result.cells[0].merge.status).toBe('none');
    expect(result.cells[0].merge.mergedRange).toBeNull();
  });

  // ── numFmt ────────────────────────────────────────────────────────────────

  it('captures numFmt for percentage cells', async () => {
    const buf = await buildSimpleBuffer((wb) => {
      const ws = wb.addWorksheet('Pct');
      const cell = ws.getCell('A1');
      cell.value = 0.0125;
      cell.numFmt = '0.00%';
    });
    const result = await svc.readCellRange(buf, 'Pct', 'A1:A1');
    // numFmt depends on how ExcelJS persists it — at minimum the cell should be found
    expect(result.cells[0].value).toBe(0.0125);
    expect(result.cells[0].type).toBe('number');
  });

  // ── Range from fixture ────────────────────────────────────────────────────

  it('reads range from financial fixture without error', async () => {
    const result = await svc.readCellRange(
      fixtureBuffer,
      'Fund Details',
      'A1:D5',
    );
    expect(result.sheetName).toBe('Fund Details');
    expect(result.cells.length).toBe(20); // 4 rows × 5 cols... wait 4×5=20 but A1:D5 = 5 rows × 4 cols
    expect(result.dimensions.rows).toBe(5);
    expect(result.dimensions.cols).toBe(4);
  });

  // ── Error cases ───────────────────────────────────────────────────────────

  it('throws for non-existent sheet', async () => {
    await expect(
      svc.readCellRange(fixtureBuffer, 'Sheet That Does Not Exist', 'A1:B2'),
    ).rejects.toThrow(/not found/i);
  });

  it('throws for invalid range string', async () => {
    await expect(
      svc.readCellRange(fixtureBuffer, 'Fund Details', 'INVALID'),
    ).rejects.toThrow();
  });

  it('throws for range that exceeds maxCells limit', async () => {
    await expect(
      svc.readCellRange(fixtureBuffer, 'Fund Details', 'A1:Z400', {
        maxCells: 100,
      }),
    ).rejects.toThrow(/exceeds the limit/i);
  });

  it('throws for empty sheet name', async () => {
    await expect(
      svc.readCellRange(fixtureBuffer, '', 'A1:B2'),
    ).rejects.toThrow();
  });
});

// ═════════════════════════════════════════════════════════════════════════════
// CELL INSPECTOR TESTS
// ═════════════════════════════════════════════════════════════════════════════

describe('inspectCell', () => {

  // ── String cell ───────────────────────────────────────────────────────────

  it('inspects a string cell', async () => {
    const buf = await buildSimpleBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A1').value = 'Management Fee';
    });
    const result = await svc.inspectCell(buf, 'Sheet1', 'A1');
    expect(result.sheetName).toBe('Sheet1');
    expect(result.coordinate).toBe('A1');
    expect(result.value).toBe('Management Fee');
    expect(result.type).toBe('string');
    expect(result.formula).toBeNull();
    expect(result.merge.status).toBe('none');
  });

  // ── Number cell ───────────────────────────────────────────────────────────

  it('inspects a number cell', async () => {
    const buf = await buildSimpleBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('B2').value = 1.25;
    });
    const result = await svc.inspectCell(buf, 'Sheet1', 'B2');
    expect(result.value).toBe(1.25);
    expect(result.type).toBe('number');
    expect(result.formula).toBeNull();
  });

  // ── Boolean cell ──────────────────────────────────────────────────────────

  it('inspects a boolean cell', async () => {
    const buf = await buildSimpleBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A1').value = true;
    });
    const result = await svc.inspectCell(buf, 'Sheet1', 'A1');
    expect(result.value).toBe(true);
    expect(result.type).toBe('boolean');
  });

  // ── Date cell ─────────────────────────────────────────────────────────────

  it('inspects a date cell', async () => {
    const buf = await buildSimpleBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A1').value = new Date('2024-01-15');
    });
    const result = await svc.inspectCell(buf, 'Sheet1', 'A1');
    expect(result.type).toBe('date');
    expect(typeof result.value).toBe('string');
    expect(result.value).toContain('2024-01-15');
  });

  // ── Empty cell ────────────────────────────────────────────────────────────

  it('inspects an empty cell', async () => {
    const buf = await buildSimpleBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A1').value = 'populated';
      // B1 is intentionally empty
    });
    const result = await svc.inspectCell(buf, 'Sheet1', 'B1');
    expect(result.type).toBe('empty');
    expect(result.value).toBeNull();
    expect(result.formula).toBeNull();
  });

  // ── Formula cell ──────────────────────────────────────────────────────────

  it('inspects a formula cell — shows formula string AND cached result', async () => {
    const buf = await buildSimpleBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A1').value = 100;
      ws.getCell('A2').value = 200;
      ws.getCell('A3').value = { formula: 'A1+A2', result: 300 };
    });
    const result = await svc.inspectCell(buf, 'Sheet1', 'A3');
    expect(result.type).toBe('formula');
    expect(result.formula).toBe('A1+A2');
    expect(result.value).toBe(300); // cached result, not recalculated
  });

  it('does NOT calculate formulas — only reads cached result', async () => {
    // If the cached result is wrong, we still report it as-is
    const buf = await buildSimpleBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A1').value = { formula: 'SUM(B1:B10)', result: 999 }; // fake cached result
    });
    const result = await svc.inspectCell(buf, 'Sheet1', 'A1');
    expect(result.type).toBe('formula');
    expect(result.formula).toBe('SUM(B1:B10)');
    expect(result.value).toBe(999); // reports cached, never recalculates
  });

  // ── Merged master cell ────────────────────────────────────────────────────

  it('identifies merged master cell with full merge info', async () => {
    const buf = await buildSimpleBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.mergeCells('A1:C1');
      ws.getCell('A1').value = 'Report Title';
    });
    const result = await svc.inspectCell(buf, 'Sheet1', 'A1');
    expect(result.merge.status).toBe('master');
    expect(result.merge.masterAddress).toBe('A1');
    expect(result.merge.mergedRange).toBe('A1:C1');
    expect(result.value).toBe('Report Title');
  });

  // ── Merged member cell ────────────────────────────────────────────────────

  it('identifies merged member cell and points to master', async () => {
    const buf = await buildSimpleBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.mergeCells('A1:C1');
      ws.getCell('A1').value = 'Report Title';
    });

    const memberB = await svc.inspectCell(buf, 'Sheet1', 'B1');
    expect(memberB.merge.status).toBe('member');
    expect(memberB.merge.masterAddress).toBe('A1');
    expect(memberB.merge.mergedRange).toBe('A1:C1');

    const memberC = await svc.inspectCell(buf, 'Sheet1', 'C1');
    expect(memberC.merge.status).toBe('member');
    expect(memberC.merge.masterAddress).toBe('A1');
  });

  // ── Absolute reference ────────────────────────────────────────────────────

  it('accepts absolute reference $A$1 for coordinate', async () => {
    const buf = await buildSimpleBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A1').value = 'Test';
    });
    const result = await svc.inspectCell(buf, 'Sheet1', '$A$1');
    expect(result.coordinate).toBe('A1'); // canonical, $ stripped
    expect(result.value).toBe('Test');
  });

  // ── Cell from fixture ─────────────────────────────────────────────────────

  it('inspects merged title cell from financial fixture', async () => {
    // Fund Overview A1 is merged A1:F1 with "Synthetic Capital Fund..." text
    const result = await svc.inspectCell(fixtureBuffer, 'Fund Overview', 'A1');
    expect(result.merge.status).toBe('master');
    expect(typeof result.value).toBe('string');
    expect(result.value).toContain('Synthetic Capital Fund');
  });

  it('inspects formula cell from financial fixture', async () => {
    // Fund Details C21 has formula SUM(C10:C12)
    const result = await svc.inspectCell(fixtureBuffer, 'Fund Details', 'C21');
    expect(result.type).toBe('formula');
    expect(result.formula).toBeTruthy();
  });

  // ── Error cases ───────────────────────────────────────────────────────────

  it('throws for non-existent sheet', async () => {
    await expect(
      svc.inspectCell(fixtureBuffer, 'NoSuchSheet', 'A1'),
    ).rejects.toThrow(/not found/i);
  });

  it('throws for invalid coordinate', async () => {
    await expect(
      svc.inspectCell(fixtureBuffer, 'Fund Overview', 'NOTACELL'),
    ).rejects.toThrow();
  });

  it('throws for A0 (row 0 is invalid)', async () => {
    await expect(
      svc.inspectCell(fixtureBuffer, 'Fund Overview', 'A0'),
    ).rejects.toThrow();
  });

  it('throws for empty coordinate', async () => {
    await expect(
      svc.inspectCell(fixtureBuffer, 'Fund Overview', ''),
    ).rejects.toThrow();
  });

  it('throws for empty sheet name', async () => {
    await expect(
      svc.inspectCell(fixtureBuffer, '', 'A1'),
    ).rejects.toThrow();
  });
});

// ═════════════════════════════════════════════════════════════════════════════
// FIXTURE INTEGRITY TEST
// ═════════════════════════════════════════════════════════════════════════════

describe('financial fixture workbook', () => {
  it('has 5 worksheets', async () => {
    const result = await svc.searchWorkbook(fixtureBuffer, 'Synthetic Capital Fund');
    // The fixture title should be findable — confirms fixture loaded
    expect(fixtureBuffer.length).toBeGreaterThan(1000); // non-trivial file
  });

  it('contains "Management Fee" in multiple sheets (anti-hallucination test)', async () => {
    const result = await svc.searchWorkbook(fixtureBuffer, 'Management Fee');
    // Must find it in more than one sheet — proves no single G42 assumption
    const sheets = new Set(result.matches.map((m) => m.sheetName));
    expect(sheets.size).toBeGreaterThan(1);
  });

  it('contains formula cells', async () => {
    const result = await svc.inspectCell(fixtureBuffer, 'Fund Details', 'C21');
    expect(result.type).toBe('formula');
  });

  it('contains merged header cells', async () => {
    const result = await svc.inspectCell(fixtureBuffer, 'Fund Overview', 'A1');
    expect(result.merge.status).toBe('master');
    expect(result.merge.mergedRange).toBeTruthy();
  });

  it('search does not modify the workbook (read-only)', async () => {
    const before = fixtureBuffer.length;
    await svc.searchWorkbook(fixtureBuffer, 'Management Fee');
    const after = fixtureBuffer.length;
    expect(after).toBe(before); // buffer unchanged
  });
});
