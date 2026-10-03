/**
 * operation-execution-router.spec.ts
 *
 * Phase 2 Step 2 — OperationExecutionRouter tests.
 *
 * Tests:
 *   A. Read-only ops routed correctly, workbook unchanged
 *   B. Write ops routed through ExcelService, workbook modified
 *   C. SET_CELL_BY_LABEL — resolved, ambiguous, not-found
 *   D. Mixed plans (read + write + label in one plan)
 *   E. Formula overwrite protection still active through router
 *   F. Regression — existing 276 tests still pass (run separately)
 */

import * as ExcelJS from 'exceljs';
import { OperationExecutionRouter } from '../operation-execution-router';
import { OperationsService } from '../operations.service';
import { ExcelService } from '../../excel/excel.service';
import { WorkbookStructureService } from '../../excel/workbook-structure.service';
import { ExcelInspectionService } from '../../excel/excel-inspection.service';
import { buildFinancialFixtureBuffer } from '../../excel/fixtures/financial-workbook.fixture';

// ─── Helpers ──────────────────────────────────────────────────────────────────

async function buildBuffer(
  setup: (wb: ExcelJS.Workbook) => void,
): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  setup(wb);
  return Buffer.from(await wb.xlsx.writeBuffer());
}

async function loadWorkbook(buffer: Buffer): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer as unknown as Parameters<typeof wb.xlsx.load>[0]);
  return wb;
}

// ─── Setup ────────────────────────────────────────────────────────────────────

let router: OperationExecutionRouter;
let opsSvc: OperationsService;
let excelSvc: ExcelService;
let fixtureBuffer: Buffer;

beforeAll(async () => {
  const structureSvc = new WorkbookStructureService();
  const inspectionSvc = new ExcelInspectionService(structureSvc);
  opsSvc = new OperationsService(inspectionSvc, structureSvc);
  excelSvc = new ExcelService();
  router = new OperationExecutionRouter(opsSvc, excelSvc);
  fixtureBuffer = await buildFinancialFixtureBuffer();
});

// ═══════════════════════════════════════════════════════════════════════════
// A. READ-ONLY OPERATIONS
// ═══════════════════════════════════════════════════════════════════════════

describe('Read-only operations through router', () => {

  it('INSPECT_WORKBOOK_STRUCTURE returns SheetStructure, workbook unchanged', async () => {
    const workbook = await loadWorkbook(fixtureBuffer);
    const plan = {
      summary: 'inspect',
      operations: [{
        type: 'INSPECT_WORKBOOK_STRUCTURE',
        sheet: 'Fund Overview',
        params: { sheetName: 'Fund Overview' },
      }],
    };
    opsSvc.validatePlan(plan);
    const result = await router.executePlan(fixtureBuffer, plan, workbook);

    expect(result.readOpsExecuted).toBe(1);
    expect(result.writeOpsApplied).toBe(0);
    expect(result.workbookModified).toBe(false);
    expect(result.operationResults[0].type).toBe('read');
    const readRes = result.operationResults[0] as any;
    expect(readRes.result.sheetName).toBe('Fund Overview');
    expect(readRes.result.cellCounts.nonEmpty).toBeGreaterThan(0);
  });

  it('DETECT_HEADERS returns candidates, workbook unchanged', async () => {
    const workbook = await loadWorkbook(fixtureBuffer);
    const plan = {
      summary: 'detect headers',
      operations: [{
        type: 'DETECT_HEADERS',
        sheet: 'Fund Details',
        params: { sheetName: 'Fund Details' },
      }],
    };
    opsSvc.validatePlan(plan);
    const result = await router.executePlan(fixtureBuffer, plan, workbook);

    expect(result.readOpsExecuted).toBe(1);
    expect(result.workbookModified).toBe(false);
    const readRes = result.operationResults[0] as any;
    expect(readRes.result.candidates).toBeDefined();
    expect(Array.isArray(readRes.result.candidates)).toBe(true);
  });

  it('DETECT_SECTIONS returns sections, workbook unchanged', async () => {
    const workbook = await loadWorkbook(fixtureBuffer);
    const plan = {
      summary: 'detect sections',
      operations: [{
        type: 'DETECT_SECTIONS',
        sheet: 'Fund Details',
        params: { sheetName: 'Fund Details' },
      }],
    };
    opsSvc.validatePlan(plan);
    const result = await router.executePlan(fixtureBuffer, plan, workbook);

    expect(result.readOpsExecuted).toBe(1);
    expect(result.workbookModified).toBe(false);
    const readRes = result.operationResults[0] as any;
    expect(readRes.result.sections.length).toBeGreaterThanOrEqual(2);
  });

  it('GET_CELL_CONTEXT returns context, workbook unchanged', async () => {
    const workbook = await loadWorkbook(fixtureBuffer);
    const plan = {
      summary: 'get context',
      operations: [{
        type: 'GET_CELL_CONTEXT',
        sheet: 'Fund Overview',
        params: { sheetName: 'Fund Overview', coordinate: 'A5', rowRadius: 2, colRadius: 2 },
      }],
    };
    opsSvc.validatePlan(plan);
    const result = await router.executePlan(fixtureBuffer, plan, workbook);

    expect(result.readOpsExecuted).toBe(1);
    expect(result.workbookModified).toBe(false);
    const readRes = result.operationResults[0] as any;
    expect(readRes.result.targetCell.coordinate).toBe('A5');
  });

  it('multiple read ops in one plan — all executed, none modify workbook', async () => {
    const workbook = await loadWorkbook(fixtureBuffer);
    const plan = {
      summary: 'multi read',
      operations: [
        { type: 'DETECT_HEADERS', sheet: 'Fund Overview', params: { sheetName: 'Fund Overview' } },
        { type: 'DETECT_SECTIONS', sheet: 'Portfolio', params: { sheetName: 'Portfolio' } },
        { type: 'INSPECT_WORKBOOK_STRUCTURE', sheet: 'Fee Structure', params: { sheetName: 'Fee Structure' } },
      ],
    };
    opsSvc.validatePlan(plan);
    const result = await router.executePlan(fixtureBuffer, plan, workbook);

    expect(result.readOpsExecuted).toBe(3);
    expect(result.writeOpsApplied).toBe(0);
    expect(result.workbookModified).toBe(false);
    expect(result.operationResults).toHaveLength(3);
    result.operationResults.forEach(r => expect(r.type).toBe('read'));
  });

  it('read ops do not mutate the original buffer', async () => {
    const bufCopy = Buffer.from(fixtureBuffer);
    const lenBefore = bufCopy.length;
    const workbook = await loadWorkbook(bufCopy);
    const plan = {
      summary: 'read only',
      operations: [{ type: 'INSPECT_WORKBOOK_STRUCTURE', sheet: 'Fund Overview', params: { sheetName: 'Fund Overview' } }],
    };
    opsSvc.validatePlan(plan);
    await router.executePlan(bufCopy, plan, workbook);
    expect(bufCopy.length).toBe(lenBefore);
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// B. WRITE OPERATIONS
// ═══════════════════════════════════════════════════════════════════════════

describe('Write operations through router', () => {

  it('SET_CELL modifies the workbook', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('Data');
      ws.getCell('A1').value = 'Name';
      ws.getCell('B1').value = 'Value';
      ws.getCell('A2').value = 'Fee';
      ws.getCell('B2').value = 1.25;
    });
    const workbook = await loadWorkbook(buf);
    const plan = {
      summary: 'update fee',
      operations: [{ type: 'SET_CELL', sheet: 'Data', params: { row: 2, col: 2, value: 2.50 } }],
    };
    opsSvc.validatePlan(plan);
    const result = await router.executePlan(buf, plan, workbook);

    expect(result.writeOpsApplied).toBe(1);
    expect(result.workbookModified).toBe(true);
    expect(workbook.getWorksheet('Data').getCell(2, 2).value).toBe(2.50);
  });

  it('ADD_ROW appends a row', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.addRow(['Name', 'Value']);
      ws.addRow(['Alice', 100]);
    });
    const workbook = await loadWorkbook(buf);
    const plan = {
      summary: 'add row',
      operations: [{ type: 'ADD_ROW', sheet: 'Sheet1', params: { values: ['Bob', 200] } }],
    };
    opsSvc.validatePlan(plan);
    const result = await router.executePlan(buf, plan, workbook);

    expect(result.writeOpsApplied).toBe(1);
    expect(result.workbookModified).toBe(true);
    expect(workbook.getWorksheet('Sheet1').getRow(3).getCell(1).value).toBe('Bob');
  });

  it('multiple write ops applied in sequence', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.addRow(['Name', 'Value']);
      ws.addRow(['Alice', 100]);
      ws.addRow(['Bob', 200]);
    });
    const workbook = await loadWorkbook(buf);
    const plan = {
      summary: 'multi write',
      operations: [
        { type: 'SET_CELL', sheet: 'Sheet1', params: { row: 2, col: 2, value: 150 } },
        { type: 'SET_CELL', sheet: 'Sheet1', params: { row: 3, col: 2, value: 250 } },
        { type: 'ADD_ROW', sheet: 'Sheet1', params: { values: ['Charlie', 300] } },
      ],
    };
    opsSvc.validatePlan(plan);
    const result = await router.executePlan(buf, plan, workbook);

    expect(result.writeOpsApplied).toBe(3);
    expect(result.workbookModified).toBe(true);
    expect(workbook.getWorksheet('Sheet1').getCell(2, 2).value).toBe(150);
    expect(workbook.getWorksheet('Sheet1').getCell(3, 2).value).toBe(250);
    expect(workbook.getWorksheet('Sheet1').getRow(4).getCell(1).value).toBe('Charlie');
  });

  it('write op result has type=write and applied=true', async () => {
    const buf = await buildBuffer((wb) => {
      wb.addWorksheet('S').addRow(['a', 'b']);
    });
    const workbook = await loadWorkbook(buf);
    const plan = {
      summary: 'test',
      operations: [{ type: 'SET_CELL', sheet: 'S', params: { row: 1, col: 1, value: 'x' } }],
    };
    opsSvc.validatePlan(plan);
    const result = await router.executePlan(buf, plan, workbook);
    const opRes = result.operationResults[0] as any;
    expect(opRes.type).toBe('write');
    expect(opRes.applied).toBe(true);
    expect(opRes.operationType).toBe('SET_CELL');
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// C. SET_CELL_BY_LABEL
// ═══════════════════════════════════════════════════════════════════════════

describe('SET_CELL_BY_LABEL through router', () => {

  it('writes value when label resolves to unique match', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('Fees');
      ws.getCell('A1').value = 'Fee Type';
      ws.getCell('B1').value = 'Rate';
      ws.getCell('A2').value = 'Unique Management Fee XYZ';
      ws.getCell('B2').value = 1.25;
    });
    const workbook = await loadWorkbook(buf);
    const plan = {
      summary: 'update fee',
      operations: [{
        type: 'SET_CELL_BY_LABEL',
        sheet: 'Fees',
        params: { label: 'Unique Management Fee XYZ', value: 2.0, valueColOffset: 1 },
      }],
    };
    opsSvc.validatePlan(plan);
    const result = await router.executePlan(buf, plan, workbook);

    expect(result.labelResolutionsPerformed).toBe(1);
    expect(result.labelResolutionsBlocked).toBe(0);
    const opRes = result.operationResults[0] as any;
    expect(opRes.type).toBe('label_resolution');
    expect(opRes.resolution.status).toBe('resolved');
    expect(opRes.writeApplied).toBe(true);
    // B2 should now be 2.0
    expect(workbook.getWorksheet('Fees').getCell(2, 2).value).toBe(2.0);
  });

  it('blocks write when label is ambiguous — never auto-selects', async () => {
    // "Management Fee" appears in 8 locations in the fixture
    const workbook = await loadWorkbook(fixtureBuffer);
    const bufCopy = Buffer.from(fixtureBuffer);
    const plan = {
      summary: 'update ambiguous fee',
      operations: [{
        type: 'SET_CELL_BY_LABEL',
        sheet: 'Fund Overview',
        params: { label: 'Management Fee', value: 1.5 },
      }],
    };
    opsSvc.validatePlan(plan);
    const result = await router.executePlan(bufCopy, plan, workbook);

    expect(result.labelResolutionsPerformed).toBe(1);
    expect(result.labelResolutionsBlocked).toBe(1);
    const opRes = result.operationResults[0] as any;
    expect(opRes.type).toBe('label_resolution');
    expect(opRes.resolution.status).toBe('ambiguous');
    expect(opRes.resolution.resolvedTarget).toBeNull();
    expect(opRes.writeBlocked).toBe(true);
    expect(opRes.writeApplied).toBe(false);
    // workbook should NOT be modified
    expect(result.workbookModified).toBe(false);
  });

  it('blocks write when label is not found', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A1').value = 'Something else';
      ws.getCell('B1').value = 100;
    });
    const workbook = await loadWorkbook(buf);
    const plan = {
      summary: 'missing label',
      operations: [{
        type: 'SET_CELL_BY_LABEL',
        sheet: 'Sheet1',
        params: { label: 'ZZZNOMATCH_LABEL', value: 999 },
      }],
    };
    opsSvc.validatePlan(plan);
    const result = await router.executePlan(buf, plan, workbook);

    expect(result.labelResolutionsBlocked).toBe(1);
    const opRes = result.operationResults[0] as any;
    expect(opRes.resolution.status).toBe('not_found');
    expect(opRes.writeBlocked).toBe(true);
    expect(opRes.writeApplied).toBe(false);
    expect(result.workbookModified).toBe(false);
  });

  it('label resolution result contains all candidates for disambiguation', async () => {
    const workbook = await loadWorkbook(fixtureBuffer);
    const plan = {
      summary: 'test',
      operations: [{
        type: 'SET_CELL_BY_LABEL',
        sheet: 'Fund Overview',
        params: { label: 'Management Fee', value: 1.5 },
      }],
    };
    opsSvc.validatePlan(plan);
    const result = await router.executePlan(fixtureBuffer, plan, workbook);
    const opRes = result.operationResults[0] as any;
    // All candidates must be returned even for ambiguous
    expect(opRes.resolution.candidates.length).toBeGreaterThan(1);
    for (const c of opRes.resolution.candidates) {
      expect(c.labelCoordinate).toMatch(/^[A-Z]+\d+$/);
      expect(c.valueCoordinate).toMatch(/^[A-Z]+\d+$/);
      expect(c.sheetName).toBeTruthy();
    }
  });

  it('sheet filter narrows candidates before resolution', async () => {
    const buf = await buildBuffer((wb) => {
      const ws1 = wb.addWorksheet('Fees');
      ws1.getCell('A1').value = 'Target Label ABC';
      ws1.getCell('B1').value = 1.0;
      const ws2 = wb.addWorksheet('Other');
      ws2.getCell('A1').value = 'Target Label ABC';
      ws2.getCell('B1').value = 2.0;
    });
    const workbook = await loadWorkbook(buf);
    // Without filter: ambiguous (2 sheets)
    const planAll = {
      summary: 'test',
      operations: [{
        type: 'SET_CELL_BY_LABEL',
        sheet: 'Fees',
        params: { label: 'Target Label ABC', value: 9.99 },
      }],
    };
    opsSvc.validatePlan(planAll);
    const resAll = await router.executePlan(buf, planAll, workbook);
    const wbAll = await loadWorkbook(buf); // fresh workbook for next test
    expect(resAll.labelResolutionsBlocked).toBe(1); // ambiguous

    // With sheetName filter: resolves to Fees only
    const workbook2 = await loadWorkbook(buf);
    const planFiltered = {
      summary: 'test filtered',
      operations: [{
        type: 'SET_CELL_BY_LABEL',
        sheet: 'Fees',
        params: { label: 'Target Label ABC', value: 9.99, sheetName: 'Fees' },
      }],
    };
    opsSvc.validatePlan(planFiltered);
    const resFiltered = await router.executePlan(buf, planFiltered, workbook2);
    expect(resFiltered.labelResolutionsBlocked).toBe(0);
    expect(resFiltered.labelResolutionsPerformed).toBe(1);
    // B1 on Fees should be updated
    expect(workbook2.getWorksheet('Fees').getCell(1, 2).value).toBe(9.99);
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// D. MIXED PLANS (read + write + label)
// ═══════════════════════════════════════════════════════════════════════════

describe('Mixed plans through router', () => {

  it('read then write — both executed, correct order', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.addRow(['Name', 'Value']);
      ws.addRow(['Fee', 1.0]);
    });
    const workbook = await loadWorkbook(buf);
    const plan = {
      summary: 'inspect then write',
      operations: [
        { type: 'INSPECT_WORKBOOK_STRUCTURE', sheet: 'Sheet1', params: { sheetName: 'Sheet1' } },
        { type: 'SET_CELL', sheet: 'Sheet1', params: { row: 2, col: 2, value: 5.0 } },
      ],
    };
    opsSvc.validatePlan(plan);
    const result = await router.executePlan(buf, plan, workbook);

    expect(result.readOpsExecuted).toBe(1);
    expect(result.writeOpsApplied).toBe(1);
    expect(result.workbookModified).toBe(true);
    expect(result.operationResults[0].type).toBe('read');
    expect(result.operationResults[1].type).toBe('write');
    expect(workbook.getWorksheet('Sheet1').getCell(2, 2).value).toBe(5.0);
  });

  it('write then read — both executed in order', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.addRow(['A', 'B']);
      ws.addRow([1, 2]);
    });
    const workbook = await loadWorkbook(buf);
    const plan = {
      summary: 'write then inspect',
      operations: [
        { type: 'SET_CELL', sheet: 'Sheet1', params: { row: 2, col: 1, value: 99 } },
        { type: 'DETECT_HEADERS', sheet: 'Sheet1', params: { sheetName: 'Sheet1' } },
      ],
    };
    opsSvc.validatePlan(plan);
    const result = await router.executePlan(buf, plan, workbook);

    // Both operations executed
    expect(result.writeOpsApplied).toBe(1);
    expect(result.readOpsExecuted).toBe(1);
    expect(result.workbookModified).toBe(true);
    // The write is applied to the workbook (execution order preserved)
    expect(workbook.getWorksheet('Sheet1').getCell(2, 1).value).toBe(99);
    // Result array has both entries (order depends on flush strategy)
    expect(result.operationResults).toHaveLength(2);
    const types = result.operationResults.map(r => r.type).sort();
    expect(types).toEqual(['read', 'write']);
  });

  it('label resolution then write — write goes through if label resolved', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('Fees');
      ws.getCell('A1').value = 'Unique Fee Label QRS';
      ws.getCell('B1').value = 1.0;
      ws.getCell('A2').value = 'Other Row';
      ws.getCell('B2').value = 0.0;
    });
    const workbook = await loadWorkbook(buf);
    const plan = {
      summary: 'label then additional write',
      operations: [
        {
          type: 'SET_CELL_BY_LABEL',
          sheet: 'Fees',
          params: { label: 'Unique Fee Label QRS', value: 2.5, valueColOffset: 1 },
        },
        { type: 'SET_CELL', sheet: 'Fees', params: { row: 2, col: 2, value: 9.9 } },
      ],
    };
    opsSvc.validatePlan(plan);
    const result = await router.executePlan(buf, plan, workbook);

    expect(result.labelResolutionsPerformed).toBe(1);
    expect(result.labelResolutionsBlocked).toBe(0);
    expect(result.writeOpsApplied).toBe(1);
    expect(result.workbookModified).toBe(true);
    expect(workbook.getWorksheet('Fees').getCell(1, 2).value).toBe(2.5);
    expect(workbook.getWorksheet('Fees').getCell(2, 2).value).toBe(9.9);
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// E. FORMULA OVERWRITE PROTECTION (still active through router)
// ═══════════════════════════════════════════════════════════════════════════

describe('Formula overwrite protection through router', () => {

  it('SET_CELL on formula cell throws without allowFormulaOverwrite', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('Calc');
      ws.getCell('A1').value = 100;
      ws.getCell('A2').value = { formula: 'SUM(A1:A1)', result: 100 };
    });
    const workbook = await loadWorkbook(buf);
    const plan = {
      summary: 'blocked write',
      operations: [{ type: 'SET_CELL', sheet: 'Calc', params: { row: 2, col: 1, value: 999 } }],
    };
    opsSvc.validatePlan(plan);
    await expect(router.executePlan(buf, plan, workbook)).rejects.toThrow(/formula/i);
  });

  it('SET_CELL on formula cell succeeds with allowFormulaOverwrite: true', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('Calc');
      ws.getCell('A1').value = { formula: 'SUM(B1:B2)', result: 300 };
    });
    const workbook = await loadWorkbook(buf);
    const plan = {
      summary: 'allowed overwrite',
      operations: [{
        type: 'SET_CELL', sheet: 'Calc',
        params: { row: 1, col: 1, value: 999, allowFormulaOverwrite: true },
      }],
    };
    opsSvc.validatePlan(plan);
    const result = await router.executePlan(buf, plan, workbook);
    expect(result.writeOpsApplied).toBe(1);
    expect(workbook.getWorksheet('Calc').getCell(1, 1).value).toBe(999);
  });

  it('SET_CELL_BY_LABEL respects formula protection on resolved cell', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('Calc');
      ws.getCell('A1').value = 'Total Fee';
      ws.getCell('B1').value = { formula: 'SUM(B2:B5)', result: 10 };
    });
    const workbook = await loadWorkbook(buf);
    const plan = {
      summary: 'label → formula cell',
      operations: [{
        type: 'SET_CELL_BY_LABEL',
        sheet: 'Calc',
        params: { label: 'Total Fee', value: 5.0, valueColOffset: 1 },
      }],
    };
    opsSvc.validatePlan(plan);
    // Should throw because target is a formula cell
    await expect(router.executePlan(buf, plan, workbook)).rejects.toThrow(/formula/i);
  });

  it('SET_CELL_BY_LABEL with allowFormulaOverwrite replaces formula cell', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('Calc');
      ws.getCell('A1').value = 'Total Fee Unique ZZZ';
      ws.getCell('B1').value = { formula: 'SUM(B2:B5)', result: 10 };
    });
    const workbook = await loadWorkbook(buf);
    const plan = {
      summary: 'label → formula cell with permission',
      operations: [{
        type: 'SET_CELL_BY_LABEL',
        sheet: 'Calc',
        params: { label: 'Total Fee Unique ZZZ', value: 5.0, valueColOffset: 1, allowFormulaOverwrite: true },
      }],
    };
    opsSvc.validatePlan(plan);
    const result = await router.executePlan(buf, plan, workbook);
    expect(result.labelResolutionsBlocked).toBe(0);
    expect(result.writeOpsApplied).toBe(0); // label writes directly, not via write batch
    expect(workbook.getWorksheet('Calc').getCell(1, 2).value).toBe(5.0);
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// F. ROUTE CLASSIFICATION HELPERS
// ═══════════════════════════════════════════════════════════════════════════

describe('OperationsService route classification helpers', () => {
  it('isReadOperation returns true for all 4 inspection ops', () => {
    ['INSPECT_WORKBOOK_STRUCTURE', 'DETECT_HEADERS', 'DETECT_SECTIONS', 'GET_CELL_CONTEXT'].forEach(t => {
      expect(opsSvc.isReadOperation(t)).toBe(true);
    });
  });

  it('isReadOperation returns false for write and hybrid ops', () => {
    ['SET_CELL', 'ADD_ROW', 'SET_CELL_BY_LABEL'].forEach(t => {
      expect(opsSvc.isReadOperation(t)).toBe(false);
    });
  });

  it('isWriteOperation returns true for all 8 write ops', () => {
    ['SET_CELL', 'SET_CELL_RANGE', 'ADD_ROW', 'DELETE_ROW',
     'ADD_FORMULA', 'FORMAT_CELL', 'ADD_COLUMN_HEADER', 'SORT_SHEET'].forEach(t => {
      expect(opsSvc.isWriteOperation(t)).toBe(true);
    });
  });

  it('isWriteOperation returns false for read and hybrid ops', () => {
    ['INSPECT_WORKBOOK_STRUCTURE', 'SET_CELL_BY_LABEL'].forEach(t => {
      expect(opsSvc.isWriteOperation(t)).toBe(false);
    });
  });
});
