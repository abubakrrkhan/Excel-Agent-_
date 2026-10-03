/**
 * phase2-step3-range-formula.spec.ts
 *
 * Phase 2 Step 3 — Formula-safe SET_CELL_RANGE tests.
 *
 * Verifies:
 *   A. SET_CELL_RANGE with no formulas — normal write succeeds
 *   B. Single formula cell in range — blocked, original intact
 *   C. Multiple formula cells in range — all detected, blocked
 *   D. Mixed range (normal + formula cells) — blocked before any write (atomic)
 *   E. Empty cells — normal write unaffected
 *   F. Single-cell SET_CELL_RANGE — consistent formula protection
 *   G. Multi-row and multi-column ranges — detection across full range
 *   H. allowFormulaOverwrite: true — explicit authorization replaces formulas
 *   I. allowFormulaOverwrite validation — must be boolean
 *   J. Error message quality — includes cell addresses, formulas, guidance
 *   K. Existing SET_CELL protection unchanged
 *   L. Existing SET_CELL_BY_LABEL protection unchanged
 *   M. OperationExecutionRouter — SET_CELL_RANGE routed correctly
 */

import * as ExcelJS from 'exceljs';
import { ExcelService } from '../../excel/excel.service';
import { OperationsService } from '../operations.service';
import { WorkbookStructureService } from '../../excel/workbook-structure.service';
import { ExcelInspectionService } from '../../excel/excel-inspection.service';
import { OperationExecutionRouter } from '../operation-execution-router';

// ─── Helpers ──────────────────────────────────────────────────────────────────

async function makeWorkbook(
  setup: (wb: ExcelJS.Workbook) => void,
): Promise<{ wb: ExcelJS.Workbook; buf: Buffer }> {
  const wb = new ExcelJS.Workbook();
  setup(wb);
  const buf = Buffer.from(await wb.xlsx.writeBuffer());
  // Reload so ExcelJS cell types resolve correctly (formula type = 6)
  const wb2 = new ExcelJS.Workbook();
  await wb2.xlsx.load(buf as unknown as Parameters<typeof wb2.xlsx.load>[0]);
  return { wb: wb2, buf };
}

// ─── Setup ────────────────────────────────────────────────────────────────────

let excelSvc: ExcelService;
let opsSvc: OperationsService;
let router: OperationExecutionRouter;

beforeAll(() => {
  excelSvc = new ExcelService();
  const structSvc = new WorkbookStructureService();
  const inspSvc = new ExcelInspectionService(structSvc);
  opsSvc = new OperationsService(inspSvc, structSvc);
  router = new OperationExecutionRouter(opsSvc, excelSvc);
});

// ═══════════════════════════════════════════════════════════════════════════
// A. No formulas — normal range write succeeds
// ═══════════════════════════════════════════════════════════════════════════

describe('A. SET_CELL_RANGE with no formula cells', () => {
  it('writes value to all cells in range when no formulas present', async () => {
    const { wb } = await makeWorkbook((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A1').value = 'x';
      ws.getCell('A2').value = 'y';
      ws.getCell('A3').value = 'z';
    });
    const plan = {
      summary: 'clear column',
      operations: [{
        type: 'SET_CELL_RANGE', sheet: 'Sheet1',
        params: { startRow: 1, startCol: 1, endRow: 3, endCol: 1, value: 0 },
      }],
    };
    opsSvc.validatePlan(plan);
    excelSvc.applyOperationPlan(wb, plan);

    const ws = wb.getWorksheet('Sheet1');
    expect(ws.getCell(1, 1).value).toBe(0);
    expect(ws.getCell(2, 1).value).toBe(0);
    expect(ws.getCell(3, 1).value).toBe(0);
  });

  it('writes across multiple columns with no formulas', async () => {
    const { wb } = await makeWorkbook((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A1').value = 1;
      ws.getCell('B1').value = 2;
      ws.getCell('A2').value = 3;
      ws.getCell('B2').value = 4;
    });
    const plan = {
      summary: 'fill range',
      operations: [{
        type: 'SET_CELL_RANGE', sheet: 'Sheet1',
        params: { startRow: 1, startCol: 1, endRow: 2, endCol: 2, value: 99 },
      }],
    };
    opsSvc.validatePlan(plan);
    excelSvc.applyOperationPlan(wb, plan);

    const ws = wb.getWorksheet('Sheet1');
    expect(ws.getCell(1, 1).value).toBe(99);
    expect(ws.getCell(1, 2).value).toBe(99);
    expect(ws.getCell(2, 1).value).toBe(99);
    expect(ws.getCell(2, 2).value).toBe(99);
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// B. Single formula cell in range — blocked
// ═══════════════════════════════════════════════════════════════════════════

describe('B. Single formula cell in range — blocked', () => {
  it('throws when range contains one formula cell (no flag)', async () => {
    const { wb } = await makeWorkbook((wb) => {
      const ws = wb.addWorksheet('Calc');
      ws.getCell('A1').value = 10;
      ws.getCell('A2').value = { formula: 'SUM(B1:B5)', result: 50 };
      ws.getCell('A3').value = 20;
    });
    const plan = {
      summary: 'test',
      operations: [{
        type: 'SET_CELL_RANGE', sheet: 'Calc',
        params: { startRow: 1, startCol: 1, endRow: 3, endCol: 1, value: 0 },
      }],
    };
    opsSvc.validatePlan(plan);
    expect(() => excelSvc.applyOperationPlan(wb, plan)).toThrow(/formula/i);
  });

  it('formula cell remains intact after blocked SET_CELL_RANGE', async () => {
    const { wb } = await makeWorkbook((wb) => {
      const ws = wb.addWorksheet('Calc');
      ws.getCell('A1').value = 10;
      ws.getCell('A2').value = { formula: 'SUM(B1:B5)', result: 50 };
      ws.getCell('A3').value = 20;
    });
    const plan = {
      summary: 'test',
      operations: [{
        type: 'SET_CELL_RANGE', sheet: 'Calc',
        params: { startRow: 1, startCol: 1, endRow: 3, endCol: 1, value: 999 },
      }],
    };
    opsSvc.validatePlan(plan);
    try { excelSvc.applyOperationPlan(wb, plan); } catch {}

    const ws = wb.getWorksheet('Calc');
    // A2 must still be a formula
    expect(ws.getCell(2, 1).type).toBe(ExcelJS.ValueType.Formula);
    const fv = ws.getCell(2, 1).value as ExcelJS.CellFormulaValue;
    expect(fv.formula).toBe('SUM(B1:B5)');
  });

  it('error message includes the protected cell address', async () => {
    const { wb } = await makeWorkbook((wb) => {
      const ws = wb.addWorksheet('Calc');
      ws.getCell('A1').value = { formula: 'A2+A3', result: 5 };
    });
    const plan = {
      summary: 'test',
      operations: [{
        type: 'SET_CELL_RANGE', sheet: 'Calc',
        params: { startRow: 1, startCol: 1, endRow: 1, endCol: 1, value: 0 },
      }],
    };
    opsSvc.validatePlan(plan);
    let errMsg = '';
    try { excelSvc.applyOperationPlan(wb, plan); } catch (e: any) { errMsg = e.message; }
    expect(errMsg).toContain('A1');
    expect(errMsg).toContain('A2+A3');
    expect(errMsg).toContain('allowFormulaOverwrite');
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// C. Multiple formula cells — all detected in one throw
// ═══════════════════════════════════════════════════════════════════════════

describe('C. Multiple formula cells in range', () => {
  it('detects all formula cells and reports them together', async () => {
    const { wb } = await makeWorkbook((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A1').value = 1;
      ws.getCell('A2').value = { formula: 'SUM(B1:B3)', result: 10 };
      ws.getCell('A3').value = 2;
      ws.getCell('A4').value = { formula: 'AVERAGE(B1:B3)', result: 5 };
      ws.getCell('A5').value = 3;
    });
    const plan = {
      summary: 'test',
      operations: [{
        type: 'SET_CELL_RANGE', sheet: 'Sheet1',
        params: { startRow: 1, startCol: 1, endRow: 5, endCol: 1, value: 0 },
      }],
    };
    opsSvc.validatePlan(plan);
    let errMsg = '';
    try { excelSvc.applyOperationPlan(wb, plan); } catch (e: any) { errMsg = e.message; }

    // Must mention both formula cells
    expect(errMsg).toContain('A2');
    expect(errMsg).toContain('A4');
    expect(errMsg).toContain('2 formula cell');
  });

  it('error message includes all formula strings', async () => {
    const { wb } = await makeWorkbook((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('B1').value = { formula: 'A1*2', result: 20 };
      ws.getCell('B2').value = { formula: 'A2+10', result: 30 };
    });
    const plan = {
      summary: 'test',
      operations: [{
        type: 'SET_CELL_RANGE', sheet: 'Sheet1',
        params: { startRow: 1, startCol: 2, endRow: 2, endCol: 2, value: 0 },
      }],
    };
    opsSvc.validatePlan(plan);
    let errMsg = '';
    try { excelSvc.applyOperationPlan(wb, plan); } catch (e: any) { errMsg = e.message; }
    expect(errMsg).toContain('A1*2');
    expect(errMsg).toContain('A2+10');
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// D. Mixed range — atomic: no partial write
// ═══════════════════════════════════════════════════════════════════════════

describe('D. Mixed range (normal + formula) — atomic, no partial write', () => {
  it('no cells are modified when a formula cell blocks the operation', async () => {
    const { wb } = await makeWorkbook((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A1').value = 'original1';
      ws.getCell('A2').value = 'original2';
      ws.getCell('A3').value = { formula: 'SUM(B1:B3)', result: 100 };
      ws.getCell('A4').value = 'original4';
      ws.getCell('A5').value = 'original5';
    });
    const plan = {
      summary: 'test',
      operations: [{
        type: 'SET_CELL_RANGE', sheet: 'Sheet1',
        params: { startRow: 1, startCol: 1, endRow: 5, endCol: 1, value: 'CHANGED' },
      }],
    };
    opsSvc.validatePlan(plan);
    try { excelSvc.applyOperationPlan(wb, plan); } catch {}

    const ws = wb.getWorksheet('Sheet1');
    // ALL cells must be unchanged — no partial write
    expect(ws.getCell(1, 1).value).toBe('original1');
    expect(ws.getCell(2, 1).value).toBe('original2');
    expect(ws.getCell(3, 1).type).toBe(ExcelJS.ValueType.Formula);
    expect(ws.getCell(4, 1).value).toBe('original4');
    expect(ws.getCell(5, 1).value).toBe('original5');
  });

  it('formula cell before normal cells — scan still finds it before any write', async () => {
    const { wb } = await makeWorkbook((wb) => {
      const ws = wb.addWorksheet('S');
      ws.getCell('A1').value = { formula: 'B1+C1', result: 5 }; // first in range
      ws.getCell('A2').value = 'normal';
      ws.getCell('A3').value = 'normal';
    });
    const plan = {
      summary: 'test',
      operations: [{
        type: 'SET_CELL_RANGE', sheet: 'S',
        params: { startRow: 1, startCol: 1, endRow: 3, endCol: 1, value: 0 },
      }],
    };
    opsSvc.validatePlan(plan);
    try { excelSvc.applyOperationPlan(wb, plan); } catch {}

    const ws = wb.getWorksheet('S');
    // A2 and A3 must NOT have been written
    expect(ws.getCell(2, 1).value).toBe('normal');
    expect(ws.getCell(3, 1).value).toBe('normal');
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// E. Empty cells — normal write unaffected
// ═══════════════════════════════════════════════════════════════════════════

describe('E. Empty cells in range', () => {
  it('writes to range with empty cells without error', async () => {
    const { wb } = await makeWorkbook((wb) => {
      const ws = wb.addWorksheet('S');
      ws.getCell('A1').value = 'filled';
      // A2, A3 are empty
      ws.getCell('A4').value = 'filled';
    });
    const plan = {
      summary: 'test',
      operations: [{
        type: 'SET_CELL_RANGE', sheet: 'S',
        params: { startRow: 1, startCol: 1, endRow: 4, endCol: 1, value: 'NEW' },
      }],
    };
    opsSvc.validatePlan(plan);
    expect(() => excelSvc.applyOperationPlan(wb, plan)).not.toThrow();

    const ws = wb.getWorksheet('S');
    expect(ws.getCell(1, 1).value).toBe('NEW');
    expect(ws.getCell(2, 1).value).toBe('NEW');
    expect(ws.getCell(3, 1).value).toBe('NEW');
    expect(ws.getCell(4, 1).value).toBe('NEW');
  });

  it('entirely empty range writes successfully', async () => {
    const { wb } = await makeWorkbook((wb) => {
      wb.addWorksheet('Empty');
    });
    const plan = {
      summary: 'test',
      operations: [{
        type: 'SET_CELL_RANGE', sheet: 'Empty',
        params: { startRow: 1, startCol: 1, endRow: 3, endCol: 3, value: 42 },
      }],
    };
    opsSvc.validatePlan(plan);
    expect(() => excelSvc.applyOperationPlan(wb, plan)).not.toThrow();
    expect(wb.getWorksheet('Empty').getCell(2, 2).value).toBe(42);
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// F. Single-cell SET_CELL_RANGE
// ═══════════════════════════════════════════════════════════════════════════

describe('F. Single-cell SET_CELL_RANGE', () => {
  it('blocks when single-cell range contains a formula', async () => {
    const { wb } = await makeWorkbook((wb) => {
      const ws = wb.addWorksheet('S');
      ws.getCell('C5').value = { formula: 'D5*2', result: 10 };
    });
    const plan = {
      summary: 'test',
      operations: [{
        type: 'SET_CELL_RANGE', sheet: 'S',
        params: { startRow: 5, startCol: 3, endRow: 5, endCol: 3, value: 0 },
      }],
    };
    opsSvc.validatePlan(plan);
    expect(() => excelSvc.applyOperationPlan(wb, plan)).toThrow(/formula/i);
    // Cell must be unchanged
    expect(wb.getWorksheet('S').getCell(5, 3).type).toBe(ExcelJS.ValueType.Formula);
  });

  it('writes when single-cell range has no formula', async () => {
    const { wb } = await makeWorkbook((wb) => {
      const ws = wb.addWorksheet('S');
      ws.getCell('C5').value = 'original';
    });
    const plan = {
      summary: 'test',
      operations: [{
        type: 'SET_CELL_RANGE', sheet: 'S',
        params: { startRow: 5, startCol: 3, endRow: 5, endCol: 3, value: 'updated' },
      }],
    };
    opsSvc.validatePlan(plan);
    expect(() => excelSvc.applyOperationPlan(wb, plan)).not.toThrow();
    expect(wb.getWorksheet('S').getCell(5, 3).value).toBe('updated');
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// G. Multi-row, multi-column range — detection across full extent
// ═══════════════════════════════════════════════════════════════════════════

describe('G. Multi-row, multi-column range', () => {
  it('detects formula in last row of multi-row range', async () => {
    const { wb } = await makeWorkbook((wb) => {
      const ws = wb.addWorksheet('S');
      ws.getCell('A1').value = 1;
      ws.getCell('A2').value = 2;
      ws.getCell('A3').value = { formula: 'A1+A2', result: 3 }; // last row
    });
    const plan = {
      summary: 'test',
      operations: [{
        type: 'SET_CELL_RANGE', sheet: 'S',
        params: { startRow: 1, startCol: 1, endRow: 3, endCol: 1, value: 0 },
      }],
    };
    opsSvc.validatePlan(plan);
    expect(() => excelSvc.applyOperationPlan(wb, plan)).toThrow(/A3/);
  });

  it('detects formula in last column of multi-column range', async () => {
    const { wb } = await makeWorkbook((wb) => {
      const ws = wb.addWorksheet('S');
      ws.getCell('A1').value = 1;
      ws.getCell('B1').value = 2;
      ws.getCell('C1').value = { formula: 'A1+B1', result: 3 }; // last col
    });
    const plan = {
      summary: 'test',
      operations: [{
        type: 'SET_CELL_RANGE', sheet: 'S',
        params: { startRow: 1, startCol: 1, endRow: 1, endCol: 3, value: 0 },
      }],
    };
    opsSvc.validatePlan(plan);
    expect(() => excelSvc.applyOperationPlan(wb, plan)).toThrow(/C1/);
  });

  it('detects formula in interior of 3x3 range', async () => {
    const { wb } = await makeWorkbook((wb) => {
      const ws = wb.addWorksheet('S');
      // Place formula at center B2
      ws.getCell('B2').value = { formula: 'SUM(A1:A3)', result: 6 };
    });
    const plan = {
      summary: 'test',
      operations: [{
        type: 'SET_CELL_RANGE', sheet: 'S',
        params: { startRow: 1, startCol: 1, endRow: 3, endCol: 3, value: 0 },
      }],
    };
    opsSvc.validatePlan(plan);
    expect(() => excelSvc.applyOperationPlan(wb, plan)).toThrow(/B2/);
  });

  it('writes successfully across 3x3 range with no formulas', async () => {
    const { wb } = await makeWorkbook((wb) => {
      const ws = wb.addWorksheet('S');
      for (let r = 1; r <= 3; r++)
        for (let c = 1; c <= 3; c++)
          ws.getCell(r, c).value = r * c;
    });
    const plan = {
      summary: 'test',
      operations: [{
        type: 'SET_CELL_RANGE', sheet: 'S',
        params: { startRow: 1, startCol: 1, endRow: 3, endCol: 3, value: 0 },
      }],
    };
    opsSvc.validatePlan(plan);
    expect(() => excelSvc.applyOperationPlan(wb, plan)).not.toThrow();
    // All 9 cells should be 0
    for (let r = 1; r <= 3; r++)
      for (let c = 1; c <= 3; c++)
        expect(wb.getWorksheet('S').getCell(r, c).value).toBe(0);
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// H. allowFormulaOverwrite: true — explicit authorization
// ═══════════════════════════════════════════════════════════════════════════

describe('H. allowFormulaOverwrite: true', () => {
  it('replaces formula cells when explicitly authorized', async () => {
    const { wb } = await makeWorkbook((wb) => {
      const ws = wb.addWorksheet('Calc');
      ws.getCell('A1').value = { formula: 'SUM(B1:B5)', result: 100 };
      ws.getCell('A2').value = 'normal';
      ws.getCell('A3').value = { formula: 'AVERAGE(B1:B5)', result: 20 };
    });
    const plan = {
      summary: 'intentional overwrite',
      operations: [{
        type: 'SET_CELL_RANGE', sheet: 'Calc',
        params: { startRow: 1, startCol: 1, endRow: 3, endCol: 1, value: 0, allowFormulaOverwrite: true },
      }],
    };
    opsSvc.validatePlan(plan);
    expect(() => excelSvc.applyOperationPlan(wb, plan)).not.toThrow();

    const ws = wb.getWorksheet('Calc');
    expect(ws.getCell(1, 1).value).toBe(0);
    expect(ws.getCell(1, 1).type).not.toBe(ExcelJS.ValueType.Formula); // no longer a formula
    expect(ws.getCell(2, 1).value).toBe(0);
    expect(ws.getCell(3, 1).value).toBe(0);
    expect(ws.getCell(3, 1).type).not.toBe(ExcelJS.ValueType.Formula);
  });

  it('allowFormulaOverwrite: false explicitly still blocks', async () => {
    const { wb } = await makeWorkbook((wb) => {
      const ws = wb.addWorksheet('S');
      ws.getCell('A1').value = { formula: 'A2+1', result: 5 };
    });
    const plan = {
      summary: 'test',
      operations: [{
        type: 'SET_CELL_RANGE', sheet: 'S',
        params: { startRow: 1, startCol: 1, endRow: 1, endCol: 1, value: 0, allowFormulaOverwrite: false },
      }],
    };
    opsSvc.validatePlan(plan);
    expect(() => excelSvc.applyOperationPlan(wb, plan)).toThrow(/formula/i);
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// I. Validation — allowFormulaOverwrite must be boolean
// ═══════════════════════════════════════════════════════════════════════════

describe('I. Validation — allowFormulaOverwrite type check', () => {
  it('rejects non-boolean allowFormulaOverwrite on SET_CELL_RANGE', () => {
    expect(() =>
      opsSvc.validatePlan({
        summary: 'test',
        operations: [{
          type: 'SET_CELL_RANGE', sheet: 'S',
          params: { startRow: 1, startCol: 1, endRow: 2, endCol: 2, value: 0, allowFormulaOverwrite: 'yes' },
        }],
      }),
    ).toThrow(/allowFormulaOverwrite must be a boolean/i);
  });

  it('accepts boolean true for allowFormulaOverwrite', () => {
    expect(() =>
      opsSvc.validatePlan({
        summary: 'test',
        operations: [{
          type: 'SET_CELL_RANGE', sheet: 'S',
          params: { startRow: 1, startCol: 1, endRow: 2, endCol: 2, value: 0, allowFormulaOverwrite: true },
        }],
      }),
    ).not.toThrow();
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// J. Error message quality
// ═══════════════════════════════════════════════════════════════════════════

describe('J. Error message quality', () => {
  it('error mentions "No cells in the range were modified"', async () => {
    const { wb } = await makeWorkbook((wb) => {
      const ws = wb.addWorksheet('S');
      ws.getCell('A1').value = { formula: 'B1+1', result: 2 };
    });
    const plan = {
      summary: 'test',
      operations: [{
        type: 'SET_CELL_RANGE', sheet: 'S',
        params: { startRow: 1, startCol: 1, endRow: 1, endCol: 1, value: 0 },
      }],
    };
    opsSvc.validatePlan(plan);
    let errMsg = '';
    try { excelSvc.applyOperationPlan(wb, plan); } catch (e: any) { errMsg = e.message; }
    expect(errMsg).toContain('No cells in the range were modified');
    expect(errMsg).toContain('allowFormulaOverwrite');
  });

  it('error includes formula count', async () => {
    const { wb } = await makeWorkbook((wb) => {
      const ws = wb.addWorksheet('S');
      ws.getCell('A1').value = { formula: 'B1', result: 1 };
      ws.getCell('A2').value = { formula: 'B2', result: 2 };
      ws.getCell('A3').value = { formula: 'B3', result: 3 };
    });
    const plan = {
      summary: 'test',
      operations: [{
        type: 'SET_CELL_RANGE', sheet: 'S',
        params: { startRow: 1, startCol: 1, endRow: 3, endCol: 1, value: 0 },
      }],
    };
    opsSvc.validatePlan(plan);
    let errMsg = '';
    try { excelSvc.applyOperationPlan(wb, plan); } catch (e: any) { errMsg = e.message; }
    expect(errMsg).toContain('3 formula cell');
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// K. Existing SET_CELL protection unchanged
// ═══════════════════════════════════════════════════════════════════════════

describe('K. Existing SET_CELL protection unchanged', () => {
  it('SET_CELL still blocks formula overwrite without flag', async () => {
    const { wb } = await makeWorkbook((wb) => {
      const ws = wb.addWorksheet('S');
      ws.getCell('A1').value = { formula: 'B1+1', result: 5 };
    });
    const plan = {
      summary: 'test',
      operations: [{ type: 'SET_CELL', sheet: 'S', params: { row: 1, col: 1, value: 99 } }],
    };
    opsSvc.validatePlan(plan);
    expect(() => excelSvc.applyOperationPlan(wb, plan)).toThrow(/formula/i);
  });

  it('SET_CELL still works with allowFormulaOverwrite: true', async () => {
    const { wb } = await makeWorkbook((wb) => {
      const ws = wb.addWorksheet('S');
      ws.getCell('A1').value = { formula: 'B1+1', result: 5 };
    });
    const plan = {
      summary: 'test',
      operations: [{ type: 'SET_CELL', sheet: 'S', params: { row: 1, col: 1, value: 99, allowFormulaOverwrite: true } }],
    };
    opsSvc.validatePlan(plan);
    expect(() => excelSvc.applyOperationPlan(wb, plan)).not.toThrow();
    expect(wb.getWorksheet('S').getCell(1, 1).value).toBe(99);
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// L. Existing SET_CELL_BY_LABEL protection unchanged
// ═══════════════════════════════════════════════════════════════════════════

describe('L. Existing SET_CELL_BY_LABEL protection unchanged', () => {
  it('SET_CELL_BY_LABEL still blocks formula at resolved target', async () => {
    const { wb, buf } = await makeWorkbook((wb) => {
      const ws = wb.addWorksheet('Fees');
      ws.getCell('A1').value = 'Management Fee Unique XYZ';
      ws.getCell('B1').value = { formula: 'C1*0.01', result: 1.25 }; // formula at target
    });
    const plan = {
      summary: 'test',
      operations: [{
        type: 'SET_CELL_BY_LABEL', sheet: 'Fees',
        params: { label: 'Management Fee Unique XYZ', value: 2.0, valueColOffset: 1 },
      }],
    };
    opsSvc.validatePlan(plan);
    await expect(router.executePlan(buf, plan, wb)).rejects.toThrow(/formula/i);
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// M. OperationExecutionRouter — SET_CELL_RANGE routed correctly
// ═══════════════════════════════════════════════════════════════════════════

describe('M. OperationExecutionRouter — SET_CELL_RANGE routing', () => {
  it('routes SET_CELL_RANGE as a write op', async () => {
    const { wb, buf } = await makeWorkbook((wb) => {
      const ws = wb.addWorksheet('S');
      ws.getCell('A1').value = 1;
      ws.getCell('A2').value = 2;
    });
    const plan = {
      summary: 'range write',
      operations: [{
        type: 'SET_CELL_RANGE', sheet: 'S',
        params: { startRow: 1, startCol: 1, endRow: 2, endCol: 1, value: 0 },
      }],
    };
    opsSvc.validatePlan(plan);
    const result = await router.executePlan(buf, plan, wb);

    expect(result.writeOpsApplied).toBe(1);
    expect(result.workbookModified).toBe(true);
    expect(wb.getWorksheet('S').getCell(1, 1).value).toBe(0);
    expect(wb.getWorksheet('S').getCell(2, 1).value).toBe(0);
  });

  it('router propagates formula protection error for SET_CELL_RANGE', async () => {
    const { wb, buf } = await makeWorkbook((wb) => {
      const ws = wb.addWorksheet('S');
      ws.getCell('A1').value = { formula: 'B1+1', result: 5 };
    });
    const plan = {
      summary: 'blocked',
      operations: [{
        type: 'SET_CELL_RANGE', sheet: 'S',
        params: { startRow: 1, startCol: 1, endRow: 2, endCol: 1, value: 0 },
      }],
    };
    opsSvc.validatePlan(plan);
    await expect(router.executePlan(buf, plan, wb)).rejects.toThrow(/formula/i);
  });

  it('router applies SET_CELL_RANGE with allowFormulaOverwrite through write path', async () => {
    const { wb, buf } = await makeWorkbook((wb) => {
      const ws = wb.addWorksheet('S');
      ws.getCell('A1').value = { formula: 'B1*2', result: 10 };
      ws.getCell('A2').value = 'normal';
    });
    const plan = {
      summary: 'authorized overwrite',
      operations: [{
        type: 'SET_CELL_RANGE', sheet: 'S',
        params: { startRow: 1, startCol: 1, endRow: 2, endCol: 1, value: 99, allowFormulaOverwrite: true },
      }],
    };
    opsSvc.validatePlan(plan);
    const result = await router.executePlan(buf, plan, wb);
    expect(result.writeOpsApplied).toBe(1);
    expect(wb.getWorksheet('S').getCell(1, 1).value).toBe(99);
    expect(wb.getWorksheet('S').getCell(2, 1).value).toBe(99);
  });
});
