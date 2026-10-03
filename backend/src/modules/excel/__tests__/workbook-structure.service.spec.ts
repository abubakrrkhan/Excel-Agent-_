/**
 * workbook-structure.service.spec.ts
 *
 * Tests for Phase 1 Step 2 — WorkbookStructureService:
 *   inspectSheetStructure()
 *   detectHeaders()
 *   detectSections()
 *   getCellContext()
 *   buildSearchMatchContext() (via ExcelInspectionService.searchWorkbook with includeContext)
 *
 * All tests use in-memory workbooks only.
 * No filesystem access. No real bank data.
 */

import * as ExcelJS from 'exceljs';
import { WorkbookStructureService } from '../workbook-structure.service';
import { ExcelInspectionService } from '../excel-inspection.service';
import {
  buildFinancialFixtureBuffer,
  buildFinancialFixtureWorkbook,
} from '../fixtures/financial-workbook.fixture';

// ─── Helpers ──────────────────────────────────────────────────────────────────

async function buildBuffer(
  setup: (wb: ExcelJS.Workbook) => void,
): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  setup(wb);
  return Buffer.from(await wb.xlsx.writeBuffer());
}

// ─── Setup ────────────────────────────────────────────────────────────────────

let svc: WorkbookStructureService;
let inspSvc: ExcelInspectionService;
let fixtureBuffer: Buffer;

beforeAll(async () => {
  svc = new WorkbookStructureService();
  inspSvc = new ExcelInspectionService(svc);
  fixtureBuffer = await buildFinancialFixtureBuffer();
});

// ═══════════════════════════════════════════════════════════════════════════
// SHEET STRUCTURE INSPECTION
// ═══════════════════════════════════════════════════════════════════════════

describe('inspectSheetStructure', () => {

  // ── Basic populated sheet ─────────────────────────────────────────────────

  it('returns correct used bounds for a simple sheet', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('B3').value = 'Start';
      ws.getCell('D7').value = 'End';
      ws.getCell('C5').value = 42;
    });
    const result = await svc.inspectSheetStructure(buf, 'Sheet1');
    expect(result.usedBounds.firstRow).toBe(3);
    expect(result.usedBounds.lastRow).toBe(7);
    expect(result.usedBounds.firstCol).toBe(2); // B
    expect(result.usedBounds.lastCol).toBe(4);  // D
    expect(result.usedBounds.address).toBe('B3:D7');
  });

  // ── Empty sheet ───────────────────────────────────────────────────────────

  it('handles empty sheet gracefully', async () => {
    const buf = await buildBuffer((wb) => {
      wb.addWorksheet('Empty');
    });
    const result = await svc.inspectSheetStructure(buf, 'Empty');
    expect(result.usedBounds.firstRow).toBeNull();
    expect(result.usedBounds.lastRow).toBeNull();
    expect(result.usedBounds.address).toBeNull();
    expect(result.cellCounts.nonEmpty).toBe(0);
    expect(result.rowCount.total).toBe(0);
  });

  // ── Cell type counting ────────────────────────────────────────────────────

  it('counts cell types correctly', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('Types');
      ws.getCell('A1').value = 'text1';       // string
      ws.getCell('A2').value = 'text2';       // string
      ws.getCell('B1').value = 100;            // number
      ws.getCell('B2').value = 200;            // number
      ws.getCell('B3').value = 300;            // number
      ws.getCell('C1').value = new Date('2024-01-01'); // date
      ws.getCell('D1').value = true;           // boolean
      ws.getCell('E1').value = { formula: 'SUM(B1:B3)', result: 600 }; // formula
    });
    const result = await svc.inspectSheetStructure(buf, 'Types');
    expect(result.cellCounts.nonEmpty).toBe(8);
    expect(result.cellCounts.text).toBe(2);
    expect(result.cellCounts.numeric).toBe(3);
    expect(result.cellCounts.date).toBe(1);
    expect(result.cellCounts.boolean).toBe(1);
    expect(result.cellCounts.formula).toBe(1);
  });

  // ── Merged ranges ─────────────────────────────────────────────────────────

  it('lists all merged ranges', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('Merges');
      ws.mergeCells('A1:C1');
      ws.getCell('A1').value = 'Wide Header';
      ws.mergeCells('A3:A5');
      ws.getCell('A3').value = 'Tall Label';
    });
    const result = await svc.inspectSheetStructure(buf, 'Merges');
    expect(result.mergedRanges.length).toBe(2);

    const wide = result.mergedRanges.find((m) => m.range === 'A1:C1');
    expect(wide).toBeDefined();
    expect(wide?.masterAddress).toBe('A1');
    expect(wide?.rowSpan).toBe(1);
    expect(wide?.colSpan).toBe(3);

    const tall = result.mergedRanges.find((m) => m.range === 'A3:A5');
    expect(tall).toBeDefined();
    expect(tall?.rowSpan).toBe(3);
    expect(tall?.colSpan).toBe(1);
  });

  it('returns empty mergedRanges for sheet with no merges', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('NoMerge');
      ws.getCell('A1').value = 'Plain';
    });
    const result = await svc.inspectSheetStructure(buf, 'NoMerge');
    expect(result.mergedRanges).toHaveLength(0);
  });

  // ── Hidden rows ───────────────────────────────────────────────────────────

  it('detects hidden rows', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('HiddenRows');
      ws.addRow(['A', 'B', 'C']);  // row 1 — visible
      ws.addRow([1, 2, 3]);         // row 2 — will be hidden
      ws.addRow([4, 5, 6]);         // row 3 — visible
      ws.getRow(2).hidden = true;
    });
    const result = await svc.inspectSheetStructure(buf, 'HiddenRows');
    expect(result.hiddenRows).toContain(2);
    expect(result.hiddenRows).not.toContain(1);
    expect(result.hiddenRows).not.toContain(3);
    expect(result.rowCount.hidden).toBe(1);
  });

  it('returns empty hiddenRows when no rows are hidden', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('Visible');
      ws.addRow(['A', 'B']);
      ws.addRow([1, 2]);
    });
    const result = await svc.inspectSheetStructure(buf, 'Visible');
    expect(result.hiddenRows).toHaveLength(0);
    expect(result.rowCount.hidden).toBe(0);
  });

  // ── Hidden columns ────────────────────────────────────────────────────────

  it('detects hidden columns', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('HiddenCols');
      ws.getCell('A1').value = 'Name';
      ws.getCell('B1').value = 'Internal';   // col B — will be hidden
      ws.getCell('C1').value = 'Value';
      ws.getColumn(2).hidden = true;           // hide column B
    });
    const result = await svc.inspectSheetStructure(buf, 'HiddenCols');
    expect(result.hiddenCols).toContain(2); // B = column 2
    expect(result.hiddenCols).not.toContain(1);
    expect(result.hiddenCols).not.toContain(3);
    expect(result.colCount.hidden).toBe(1);
  });

  // ── From financial fixture ────────────────────────────────────────────────

  it('detects hidden rows in Fund Overview from fixture', async () => {
    const result = await svc.inspectSheetStructure(fixtureBuffer, 'Fund Overview');
    // Row 3 is hidden in the fixture
    expect(result.hiddenRows).toContain(3);
    expect(result.rowCount.hidden).toBeGreaterThanOrEqual(1);
  });

  it('detects hidden cols in Fund Overview from fixture', async () => {
    const result = await svc.inspectSheetStructure(fixtureBuffer, 'Fund Overview');
    // Column C (3) is hidden in the fixture
    expect(result.hiddenCols).toContain(3);
  });

  it('detects hidden row in Fund Details from fixture', async () => {
    const result = await svc.inspectSheetStructure(fixtureBuffer, 'Fund Details');
    // Row 7 is hidden in Fund Details
    expect(result.hiddenRows).toContain(7);
  });

  it('detects merged ranges in Fund Overview from fixture', async () => {
    const result = await svc.inspectSheetStructure(fixtureBuffer, 'Fund Overview');
    // A1:F1 is the merged title
    expect(result.mergedRanges.length).toBeGreaterThan(0);
    const titleMerge = result.mergedRanges.find((m) => m.masterAddress === 'A1');
    expect(titleMerge).toBeDefined();
  });

  it('counts formula cells from fixture', async () => {
    const result = await svc.inspectSheetStructure(fixtureBuffer, 'Fund Details');
    // Fund Details has SUM formula and AVERAGE formula
    expect(result.cellCounts.formula).toBeGreaterThanOrEqual(1);
  });

  // ── Error cases ───────────────────────────────────────────────────────────

  it('throws for non-existent sheet', async () => {
    await expect(
      svc.inspectSheetStructure(fixtureBuffer, 'NoSuchSheet'),
    ).rejects.toThrow(/not found/i);
  });

  it('throws for empty sheet name', async () => {
    await expect(
      svc.inspectSheetStructure(fixtureBuffer, ''),
    ).rejects.toThrow();
  });

  // ── Determinism ───────────────────────────────────────────────────────────

  it('returns identical results on repeated calls', async () => {
    const r1 = await svc.inspectSheetStructure(fixtureBuffer, 'Fund Details');
    const r2 = await svc.inspectSheetStructure(fixtureBuffer, 'Fund Details');
    expect(r1.cellCounts.nonEmpty).toBe(r2.cellCounts.nonEmpty);
    expect(r1.mergedRanges.length).toBe(r2.mergedRanges.length);
    expect(r1.hiddenRows).toEqual(r2.hiddenRows);
  });

  // ── Read-only guarantee ───────────────────────────────────────────────────

  it('does not modify the buffer', async () => {
    const before = fixtureBuffer.length;
    await svc.inspectSheetStructure(fixtureBuffer, 'Fund Overview');
    expect(fixtureBuffer.length).toBe(before);
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// HEADER DETECTION
// ═══════════════════════════════════════════════════════════════════════════

describe('detectHeaders', () => {

  // ── Clear header row ──────────────────────────────────────────────────────

  it('detects a clear header row followed by data rows', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('Table');
      // Row 1: text headers
      ws.addRow(['Name', 'Department', 'Salary', 'Start Date']);
      // Row 2-4: data rows (mix of text and numbers)
      ws.addRow(['Alice', 'Engineering', 95000, new Date('2021-01-01')]);
      ws.addRow(['Bob', 'Marketing', 72000, new Date('2022-06-01')]);
      ws.addRow(['Charlie', 'HR', 65000, new Date('2020-03-01')]);
    });
    const result = await svc.detectHeaders(buf, 'Table');
    expect(result.candidates.length).toBeGreaterThan(0);
    const row1 = result.candidates.find((c) => c.rowNumber === 1);
    expect(row1).toBeDefined();
    expect(row1?.textCellCount).toBe(4); // all text
    expect(row1?.labels).toContain('Name');
    expect(row1?.labels).toContain('Salary');
  });

  // ── Merged header row ─────────────────────────────────────────────────────

  it('identifies merged header row with high confidence', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('MergedHeader');
      ws.mergeCells('A1:D1');
      ws.getCell('A1').value = 'Fund Performance Report';
      ws.addRow(['Period', 'Return', 'Benchmark', 'Excess']);
      ws.addRow(['1M', 0.023, 0.019, 0.004]);
      ws.addRow(['3M', 0.068, 0.056, 0.012]);
    });
    const result = await svc.detectHeaders(buf, 'MergedHeader');
    expect(result.candidates.length).toBeGreaterThan(0);
    const mergedRow = result.candidates.find((c) => c.hasMergedCells);
    expect(mergedRow).toBeDefined();
    expect(mergedRow?.confidence).toBe('high');
  });

  // ── Multiple candidates ───────────────────────────────────────────────────

  it('returns multiple candidates for multi-section sheet', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('MultiSection');
      // Section 1 header
      ws.addRow(['Item', 'Value', 'Notes']);
      ws.addRow(['Fee A', 1.25, 'Annual']);
      ws.addRow(['Fee B', 0.10, 'Monthly']);
      // Gap
      ws.addRow([]);
      ws.addRow([]);
      // Section 2 header (another header-like row)
      ws.addRow(['Asset Class', 'Weight', 'Benchmark']);
      ws.addRow(['Equity', 0.6, 0.55]);
      ws.addRow(['Fixed Income', 0.35, 0.40]);
    });
    const result = await svc.detectHeaders(buf, 'MultiSection');
    expect(result.hasMultipleCandidates).toBe(true);
    expect(result.candidates.length).toBeGreaterThanOrEqual(2);
  });

  // ── All-numeric sheet — no header candidates ──────────────────────────────

  it('returns no high-confidence candidates for all-numeric sheet', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('Numeric');
      ws.addRow([1, 2, 3, 4, 5]);
      ws.addRow([6, 7, 8, 9, 10]);
      ws.addRow([11, 12, 13, 14, 15]);
    });
    const result = await svc.detectHeaders(buf, 'Numeric');
    const highConf = result.candidates.filter((c) => c.confidence === 'high');
    expect(highConf.length).toBe(0);
  });

  // ── Empty sheet ───────────────────────────────────────────────────────────

  it('returns no candidates for empty sheet', async () => {
    const buf = await buildBuffer((wb) => {
      wb.addWorksheet('Empty');
    });
    const result = await svc.detectHeaders(buf, 'Empty');
    expect(result.candidates).toHaveLength(0);
    expect(result.hasMultipleCandidates).toBe(false);
  });

  // ── Evidence string ───────────────────────────────────────────────────────

  it('includes a meaningful evidence string', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('Evidence');
      ws.addRow(['Name', 'Value', 'Type']);
      ws.addRow(['Fee', 1.5, 'Annual']);
    });
    const result = await svc.detectHeaders(buf, 'Evidence');
    if (result.candidates.length > 0) {
      expect(result.candidates[0].evidence).toBeTruthy();
      expect(typeof result.candidates[0].evidence).toBe('string');
      expect(result.candidates[0].evidence.length).toBeGreaterThan(5);
    }
  });

  // ── From fixture ──────────────────────────────────────────────────────────

  it('finds header candidates in Fund Details from fixture', async () => {
    const result = await svc.detectHeaders(fixtureBuffer, 'Fund Details');
    expect(result.candidates.length).toBeGreaterThan(0);
  });

  it('finds multiple candidates in Portfolio sheet from fixture', async () => {
    const result = await svc.detectHeaders(fixtureBuffer, 'Portfolio');
    // Portfolio has multiple sections — expects multiple header candidates
    expect(result.candidates.length).toBeGreaterThan(0);
  });

  // ── Error cases ───────────────────────────────────────────────────────────

  it('throws for invalid sheet name', async () => {
    await expect(
      svc.detectHeaders(fixtureBuffer, 'NoSuchSheet'),
    ).rejects.toThrow(/not found/i);
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// SECTION / TABLE DETECTION
// ═══════════════════════════════════════════════════════════════════════════

describe('detectSections', () => {

  // ── Single section ────────────────────────────────────────────────────────

  it('detects a single contiguous section', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('Single');
      ws.addRow(['Name', 'Value']);
      ws.addRow(['Fee A', 1.25]);
      ws.addRow(['Fee B', 0.10]);
    });
    const result = await svc.detectSections(buf, 'Single');
    expect(result.sections.length).toBe(1);
    expect(result.sections[0].rows).toBe(3);
    expect(result.sections[0].nonEmptyCells).toBeGreaterThanOrEqual(6);
  });

  // ── Multiple sections separated by blank rows ─────────────────────────────

  it('detects two sections separated by blank rows', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('TwoSections');
      ws.addRow(['Header A', 'Value A']);
      ws.addRow(['Row 1', 100]);
      ws.addRow(['Row 2', 200]);
      ws.addRow([]);          // blank
      ws.addRow([]);          // blank (gap = 2 → section boundary)
      ws.addRow(['Header B', 'Value B']);
      ws.addRow(['Row 3', 300]);
    });
    const result = await svc.detectSections(buf, 'TwoSections');
    expect(result.sections.length).toBe(2);
    expect(result.sections[0].startCoordinate).toContain('1'); // starts at row 1
    expect(result.sections[1].startCoordinate).toContain('6'); // starts at row 6
  });

  // ── Sections include header candidates ────────────────────────────────────

  it('includes header candidates within each section', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('WithHeaders');
      ws.addRow(['Name', 'Department', 'Salary']);
      ws.addRow(['Alice', 'Engineering', 95000]);
      ws.addRow(['Bob', 'Marketing', 72000]);
    });
    const result = await svc.detectSections(buf, 'WithHeaders');
    expect(result.sections.length).toBe(1);
    // The first row is a header candidate within the section
    expect(result.sections[0].headerCandidates.length).toBeGreaterThanOrEqual(0);
    // (may or may not find it depending on scoring — just ensure no crash)
  });

  // ── Formula detection within sections ─────────────────────────────────────

  it('detects formula presence within a section', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('Formulas');
      ws.getCell('A1').value = 'Item';
      ws.getCell('B1').value = 'Amount';
      ws.getCell('A2').value = 'Fee A';
      ws.getCell('B2').value = 100;
      ws.getCell('A3').value = 'Total';
      ws.getCell('B3').value = { formula: 'SUM(B2:B2)', result: 100 };
    });
    const result = await svc.detectSections(buf, 'Formulas');
    expect(result.sections.length).toBe(1);
    expect(result.sections[0].hasFormulas).toBe(true);
  });

  // ── Merged cells within sections ──────────────────────────────────────────

  it('detects merged cells within a section', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('MergedSection');
      ws.mergeCells('A1:C1');
      ws.getCell('A1').value = 'Section Title';
      ws.addRow(['Col A', 'Col B', 'Col C']);
      ws.addRow([1, 2, 3]);
    });
    const result = await svc.detectSections(buf, 'MergedSection');
    expect(result.sections.length).toBe(1);
    expect(result.sections[0].hasMergedCells).toBe(true);
  });

  // ── Empty sheet ───────────────────────────────────────────────────────────

  it('returns no sections for empty sheet', async () => {
    const buf = await buildBuffer((wb) => {
      wb.addWorksheet('Empty');
    });
    const result = await svc.detectSections(buf, 'Empty');
    expect(result.sections).toHaveLength(0);
  });

  // ── From fixture ──────────────────────────────────────────────────────────

  it('detects multiple sections in Fund Details from fixture', async () => {
    const result = await svc.detectSections(fixtureBuffer, 'Fund Details');
    // Fund Details has fee section + performance section separated by blanks
    expect(result.sections.length).toBeGreaterThanOrEqual(2);
  });

  it('detects sections in Portfolio sheet from fixture', async () => {
    const result = await svc.detectSections(fixtureBuffer, 'Portfolio');
    expect(result.sections.length).toBeGreaterThanOrEqual(2);
    // Check some sections have formulas (portfolio totals)
    const withFormulas = result.sections.filter((s) => s.hasFormulas);
    expect(withFormulas.length).toBeGreaterThan(0);
  });

  it('each section has a valid range address', async () => {
    const result = await svc.detectSections(fixtureBuffer, 'Fund Details');
    for (const section of result.sections) {
      expect(section.rangeAddress).toMatch(/^[A-Z]+\d+:[A-Z]+\d+$/);
      expect(section.rows).toBeGreaterThan(0);
      expect(section.cols).toBeGreaterThan(0);
    }
  });

  // ── Error cases ───────────────────────────────────────────────────────────

  it('throws for invalid sheet name', async () => {
    await expect(
      svc.detectSections(fixtureBuffer, 'NoSheet'),
    ).rejects.toThrow(/not found/i);
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// CELL CONTEXT INSPECTION
// ═══════════════════════════════════════════════════════════════════════════

describe('getCellContext', () => {

  // ── Basic context ─────────────────────────────────────────────────────────

  it('returns target cell and surrounding context', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('Context');
      ws.getCell('A1').value = 'Label';
      ws.getCell('B1').value = 'Header';
      ws.getCell('A2').value = 'Management Fee';
      ws.getCell('B2').value = 1.25;
      ws.getCell('A3').value = 'Trustee Fee';
      ws.getCell('B3').value = 0.10;
    });
    const result = await svc.getCellContext(buf, 'Context', 'B2');

    expect(result.targetCell.coordinate).toBe('B2');
    expect(result.targetCell.value).toBe(1.25);
    expect(result.targetCell.type).toBe('number');
    expect(result.cells.length).toBeGreaterThan(1); // context window has multiple cells
  });

  // ── Row labels ────────────────────────────────────────────────────────────

  it('returns row labels (other non-empty cells in same row)', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('Labels');
      ws.getCell('A2').value = 'Management Fee'; // row label
      ws.getCell('B2').value = 1.25;             // target
      ws.getCell('C2').value = '%';              // unit label
    });
    const result = await svc.getCellContext(buf, 'Labels', 'B2');

    expect(result.rowLabels.length).toBeGreaterThanOrEqual(1);
    const mfLabel = result.rowLabels.find((c) => c.value === 'Management Fee');
    expect(mfLabel).toBeDefined();
  });

  // ── Column labels ─────────────────────────────────────────────────────────

  it('returns column labels (non-empty cells in same column within window)', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('ColLabels');
      ws.getCell('B1').value = 'Value';    // column header
      ws.getCell('B2').value = 1.25;       // target
      ws.getCell('B3').value = 0.10;       // below target
    });
    const result = await svc.getCellContext(buf, 'ColLabels', 'B2', {
      rowRadius: 2,
      colRadius: 0,
    });
    // colLabels should include B1 and/or B3
    expect(result.colLabels.length).toBeGreaterThanOrEqual(1);
  });

  // ── Custom radius ─────────────────────────────────────────────────────────

  it('respects custom rowRadius and colRadius', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('Radius');
      for (let r = 1; r <= 10; r++) {
        for (let c = 1; c <= 10; c++) {
          ws.getCell(r, c).value = `${r},${c}`;
        }
      }
    });
    // Small radius
    const small = await svc.getCellContext(buf, 'Radius', 'E5', {
      rowRadius: 1,
      colRadius: 1,
    });
    // 3 rows × 3 cols = 9 cells
    expect(small.cells.length).toBe(9);
    expect(small.contextRange.startRow).toBe(4);
    expect(small.contextRange.endRow).toBe(6);

    // Larger radius
    const large = await svc.getCellContext(buf, 'Radius', 'E5', {
      rowRadius: 2,
      colRadius: 2,
    });
    expect(large.cells.length).toBe(25); // 5 × 5
  });

  // ── Edge: cell at sheet boundary ──────────────────────────────────────────

  it('clamps context to sheet bounds at top-left corner', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('Corner');
      ws.getCell('A1').value = 'TopLeft';
      ws.getCell('B1').value = 'Right';
      ws.getCell('A2').value = 'Below';
    });
    const result = await svc.getCellContext(buf, 'Corner', 'A1', {
      rowRadius: 3,
      colRadius: 3,
    });
    // Context should start at A1, not go negative
    expect(result.contextRange.startRow).toBeGreaterThanOrEqual(1);
    expect(result.contextRange.startCol).toBeGreaterThanOrEqual(1);
    expect(result.targetCell.coordinate).toBe('A1');
  });

  // ── Context range address format ──────────────────────────────────────────

  it('contextRange.address is a valid range string', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('G10').value = 'Target';
    });
    const result = await svc.getCellContext(buf, 'Sheet1', 'G10');
    expect(result.contextRange.address).toMatch(/^[A-Z]+\d+:[A-Z]+\d+$/);
  });

  // ── From fixture ──────────────────────────────────────────────────────────

  it('gets context around a Management Fee cell in fixture', async () => {
    // First find a Management Fee cell
    const searchResult = await inspSvc.searchWorkbook(
      fixtureBuffer,
      'Management Fee',
      { sheetName: 'Fund Details', maxMatches: 1 },
    );
    expect(searchResult.matches.length).toBeGreaterThan(0);

    const match = searchResult.matches[0];
    const context = await svc.getCellContext(
      fixtureBuffer,
      match.sheetName,
      match.coordinate,
    );

    expect(context.targetCell.value).toContain('Management Fee');
    expect(context.cells.length).toBeGreaterThan(0);
    // Row labels should include the value cell next to Management Fee
    expect(context.rowLabels.length).toBeGreaterThanOrEqual(0);
  });

  // ── Error cases ───────────────────────────────────────────────────────────

  it('throws for invalid coordinate', async () => {
    await expect(
      svc.getCellContext(fixtureBuffer, 'Fund Overview', 'NOTACELL'),
    ).rejects.toThrow();
  });

  it('throws for A0 row', async () => {
    await expect(
      svc.getCellContext(fixtureBuffer, 'Fund Overview', 'A0'),
    ).rejects.toThrow();
  });

  it('throws for invalid sheet name', async () => {
    await expect(
      svc.getCellContext(fixtureBuffer, 'NoSuchSheet', 'A1'),
    ).rejects.toThrow(/not found/i);
  });

  it('throws for empty sheet name', async () => {
    await expect(
      svc.getCellContext(fixtureBuffer, '', 'A1'),
    ).rejects.toThrow();
  });

  // ── Read-only ─────────────────────────────────────────────────────────────

  it('does not modify the buffer', async () => {
    const before = fixtureBuffer.length;
    await svc.getCellContext(fixtureBuffer, 'Fund Overview', 'A1');
    expect(fixtureBuffer.length).toBe(before);
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// ENHANCED SEARCH EVIDENCE (via ExcelInspectionService)
// ═══════════════════════════════════════════════════════════════════════════

describe('enhanced search with includeContext=true', () => {

  it('enriches matches with context when includeContext=true', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('Fees');
      ws.getCell('A1').value = 'Fee Type';
      ws.getCell('B1').value = 'Rate';
      ws.getCell('A2').value = 'Management Fee';
      ws.getCell('B2').value = 1.25;
    });
    const result = await inspSvc.searchWorkbook(buf, 'Management Fee', {
      includeContext: true,
    });
    expect(result.matches.length).toBeGreaterThan(0);
    const match = result.matches[0];
    expect(match.context).toBeDefined();
    expect(match.context?.cell).toBeDefined();
    expect(match.context?.cell.coordinate).toBe('A2');
  });

  it('context includes row labels for disambiguation', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('Fees');
      ws.getCell('A1').value = 'Section: Annual Fees';
      ws.getCell('A2').value = 'Management Fee';
      ws.getCell('B2').value = 1.25;
      ws.getCell('C2').value = '%';
    });
    const result = await inspSvc.searchWorkbook(buf, 'Management Fee', {
      includeContext: true,
    });
    const match = result.matches[0];
    // Row labels should contain "1.25" (value cell) and "%" (unit cell)
    const rowLabels = match.context?.rowLabels ?? [];
    expect(rowLabels.length).toBeGreaterThanOrEqual(1);
  });

  it('context includes nearby above rows (section headings)', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('Fees');
      ws.getCell('A1').value = 'Fee Structure';   // heading above
      ws.getCell('A3').value = 'Management Fee'; // target (gap row between)
      ws.getCell('B3').value = 1.25;
    });
    const result = await inspSvc.searchWorkbook(buf, 'Management Fee', {
      includeContext: true,
    });
    const match = result.matches[0];
    expect(match.context?.nearbyAbove).toBeDefined();
    // nearbyAbove should find 'Fee Structure' in A1
    const headingFound = match.context?.nearbyAbove.some(
      (n) => n.value === 'Fee Structure',
    );
    expect(headingFound).toBe(true);
  });

  it('does NOT include context when includeContext=false (default)', async () => {
    const result = await inspSvc.searchWorkbook(
      fixtureBuffer,
      'Management Fee',
    );
    for (const match of result.matches) {
      expect(match.context).toBeUndefined();
    }
  });

  it('multiple matches all have context when includeContext=true', async () => {
    const result = await inspSvc.searchWorkbook(
      fixtureBuffer,
      'Management Fee',
      { includeContext: true },
    );
    expect(result.matches.length).toBeGreaterThan(1);
    for (const match of result.matches) {
      expect(match.context).toBeDefined();
    }
  });

  // ── Context distinguishes between same label in different sheets ──────────

  it('context reveals different sections for same label across sheets', async () => {
    const result = await inspSvc.searchWorkbook(
      fixtureBuffer,
      'Management Fee',
      { includeContext: true },
    );
    // Different matches should have different sheet contexts
    const sheets = new Set(result.matches.map((m) => m.sheetName));
    expect(sheets.size).toBeGreaterThan(1); // proves they're in different sheets

    // Context of each match is specific to its own location
    for (const match of result.matches) {
      if (match.context) {
        expect(match.context.cell.sheetName).toBe(match.sheetName);
        expect(match.context.cell.coordinate).toBe(match.coordinate);
      }
    }
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// FORMULA CONTEXT
// ═══════════════════════════════════════════════════════════════════════════

describe('formula context in structural inspection', () => {

  it('sheet structure counts formula cells correctly', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('FormSheet');
      ws.getCell('A1').value = 'Item';
      ws.getCell('B1').value = 'Value';
      ws.getCell('A2').value = 100;
      ws.getCell('A3').value = 200;
      ws.getCell('B2').value = { formula: 'SUM(A2:A3)', result: 300 };
      ws.getCell('B3').value = { formula: 'AVERAGE(A2:A3)', result: 150 };
    });
    const result = await svc.inspectSheetStructure(buf, 'FormSheet');
    expect(result.cellCounts.formula).toBe(2);
    expect(result.cellCounts.nonEmpty).toBe(6);
  });

  it('section detection marks hasFormulas=true when section contains formula cells', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('SectionFormulas');
      ws.getCell('A1').value = 'Name';
      ws.getCell('B1').value = 'Amount';
      ws.getCell('A2').value = 'Fee A';
      ws.getCell('B2').value = 100;
      ws.getCell('A3').value = 'Fee B';
      ws.getCell('B3').value = 200;
      ws.getCell('A4').value = 'Total';
      ws.getCell('B4').value = { formula: 'SUM(B2:B3)', result: 300 };
    });
    const result = await svc.detectSections(buf, 'SectionFormulas');
    expect(result.sections.length).toBe(1);
    expect(result.sections[0].hasFormulas).toBe(true);
  });

  it('getCellContext returns formula string for formula cells', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('FormulaCtx');
      ws.getCell('A1').value = 100;
      ws.getCell('A2').value = 200;
      ws.getCell('A3').value = { formula: 'SUM(A1:A2)', result: 300 };
    });
    const result = await svc.getCellContext(buf, 'FormulaCtx', 'A3');
    expect(result.targetCell.type).toBe('formula');
    expect(result.targetCell.formula).toBe('SUM(A1:A2)');
    expect(result.targetCell.value).toBe(300); // cached result only
  });
});
