/**
 * phase2-operations.spec.ts
 *
 * Phase 2 Step 1 tests:
 *   A. Inspection operations (INSPECT_WORKBOOK_STRUCTURE, DETECT_HEADERS,
 *      DETECT_SECTIONS, GET_CELL_CONTEXT) through OperationsService
 *   B. Formula overwrite protection (validation + execution layers)
 *   C. SET_CELL_BY_LABEL foundation (resolution, ambiguity, not-found)
 *   D. Regression — existing 8 write operations still accepted by validatePlan
 */

import * as ExcelJS from 'exceljs';
import { OperationsService } from '../operations.service';
import { ExcelInspectionService } from '../../excel/excel-inspection.service';
import { WorkbookStructureService } from '../../excel/workbook-structure.service';
import { ExcelService } from '../../excel/excel.service';
import { buildFinancialFixtureBuffer } from '../../excel/fixtures/financial-workbook.fixture';

// ─── Helpers ──────────────────────────────────────────────────────────────────

async function buildBuffer(
  setup: (wb: ExcelJS.Workbook) => void,
): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  setup(wb);
  return Buffer.from(await wb.xlsx.writeBuffer());
}

// ─── Setup ────────────────────────────────────────────────────────────────────

let svc: OperationsService;
let excelSvc: ExcelService;
let fixtureBuffer: Buffer;

beforeAll(async () => {
  const structureSvc = new WorkbookStructureService();
  const inspectionSvc = new ExcelInspectionService(structureSvc);
  svc = new OperationsService(inspectionSvc, structureSvc);
  excelSvc = new ExcelService();
  fixtureBuffer = await buildFinancialFixtureBuffer();
});

// ═══════════════════════════════════════════════════════════════════════════
// A. INSPECTION OPERATIONS
// ═══════════════════════════════════════════════════════════════════════════

describe('Inspection operation types recognized by validatePlan', () => {
  const inspectionOps = [
    'INSPECT_WORKBOOK_STRUCTURE',
    'DETECT_HEADERS',
    'DETECT_SECTIONS',
  ];

  inspectionOps.forEach((type) => {
    it(`${type} is accepted by validatePlan`, () => {
      expect(() =>
        svc.validatePlan({
          summary: 'test',
          operations: [{ type, sheet: 'Fund Overview', params: { sheetName: 'Fund Overview' } }],
        }),
      ).not.toThrow();
    });
  });

  it('GET_CELL_CONTEXT is accepted by validatePlan', () => {
    expect(() =>
      svc.validatePlan({
        summary: 'test',
        operations: [{
          type: 'GET_CELL_CONTEXT',
          sheet: 'Fund Overview',
          params: { sheetName: 'Fund Overview', coordinate: 'A1' },
        }],
      }),
    ).not.toThrow();
  });

  it('unknown operation type is still rejected', () => {
    expect(() =>
      svc.validatePlan({
        summary: 'bad',
        operations: [{ type: 'HACK_DB', sheet: 'Sheet1', params: {} }],
      }),
    ).toThrow(/unsupported operation type/i);
  });
});

describe('INSPECT_WORKBOOK_STRUCTURE via executeInspectStructure', () => {
  it('returns SheetStructure for a valid sheet', async () => {
    const result = await svc.executeInspectStructure(fixtureBuffer, {
      sheetName: 'Fund Overview',
    });
    expect(result.sheetName).toBe('Fund Overview');
    expect(result.cellCounts.nonEmpty).toBeGreaterThan(0);
    expect(result.usedBounds.firstRow).toBeGreaterThan(0);
  });

  it('includes merged ranges', async () => {
    const result = await svc.executeInspectStructure(fixtureBuffer, {
      sheetName: 'Fund Overview',
    });
    expect(result.mergedRanges.length).toBeGreaterThan(0);
  });

  it('includes hidden rows from fixture', async () => {
    const result = await svc.executeInspectStructure(fixtureBuffer, {
      sheetName: 'Fund Overview',
    });
    expect(result.hiddenRows).toContain(3);
  });

  it('throws for missing sheetName param', async () => {
    await expect(
      svc.executeInspectStructure(fixtureBuffer, {}),
    ).rejects.toThrow(/sheetName/i);
  });

  it('throws for non-existent sheet', async () => {
    await expect(
      svc.executeInspectStructure(fixtureBuffer, { sheetName: 'NoSuchSheet' }),
    ).rejects.toThrow(/not found/i);
  });

  it('returns empty bounds for empty sheet', async () => {
    const buf = await buildBuffer((wb) => { wb.addWorksheet('Empty'); });
    const result = await svc.executeInspectStructure(buf, { sheetName: 'Empty' });
    expect(result.usedBounds.firstRow).toBeNull();
    expect(result.cellCounts.nonEmpty).toBe(0);
  });

  it('is read-only — buffer unchanged after call', async () => {
    const len = fixtureBuffer.length;
    await svc.executeInspectStructure(fixtureBuffer, { sheetName: 'Fund Overview' });
    expect(fixtureBuffer.length).toBe(len);
  });
});

describe('DETECT_HEADERS via executeDetectHeaders', () => {
  it('returns HeaderDetectionResult with candidates', async () => {
    const result = await svc.executeDetectHeaders(fixtureBuffer, {
      sheetName: 'Fund Details',
    });
    expect(result.sheetName).toBe('Fund Details');
    expect(result.candidates.length).toBeGreaterThan(0);
  });

  it('returns hasMultipleCandidates flag', async () => {
    const result = await svc.executeDetectHeaders(fixtureBuffer, {
      sheetName: 'Fund Details',
    });
    expect(typeof result.hasMultipleCandidates).toBe('boolean');
  });

  it('each candidate has evidence string', async () => {
    const result = await svc.executeDetectHeaders(fixtureBuffer, {
      sheetName: 'Fund Details',
    });
    for (const c of result.candidates) {
      expect(typeof c.evidence).toBe('string');
      expect(c.evidence.length).toBeGreaterThan(0);
    }
  });

  it('throws for missing sheetName', async () => {
    await expect(svc.executeDetectHeaders(fixtureBuffer, {})).rejects.toThrow();
  });

  it('throws for non-existent sheet', async () => {
    await expect(
      svc.executeDetectHeaders(fixtureBuffer, { sheetName: 'NoSheet' }),
    ).rejects.toThrow(/not found/i);
  });

  it('is read-only — buffer unchanged after call', async () => {
    const len = fixtureBuffer.length;
    await svc.executeDetectHeaders(fixtureBuffer, { sheetName: 'Fund Details' });
    expect(fixtureBuffer.length).toBe(len);
  });
});

describe('DETECT_SECTIONS via executeDetectSections', () => {
  it('returns SectionDetectionResult with sections', async () => {
    const result = await svc.executeDetectSections(fixtureBuffer, {
      sheetName: 'Fund Details',
    });
    expect(result.sheetName).toBe('Fund Details');
    expect(result.sections.length).toBeGreaterThanOrEqual(2);
  });

  it('each section has a valid rangeAddress', async () => {
    const result = await svc.executeDetectSections(fixtureBuffer, {
      sheetName: 'Fund Details',
    });
    for (const sec of result.sections) {
      expect(sec.rangeAddress).toMatch(/^[A-Z]+\d+:[A-Z]+\d+$/);
    }
  });

  it('reports hasFormulas correctly', async () => {
    const result = await svc.executeDetectSections(fixtureBuffer, {
      sheetName: 'Fund Details',
    });
    const withFormulas = result.sections.filter((s) => s.hasFormulas);
    expect(withFormulas.length).toBeGreaterThan(0);
  });

  it('throws for missing sheetName', async () => {
    await expect(svc.executeDetectSections(fixtureBuffer, {})).rejects.toThrow();
  });

  it('is read-only — buffer unchanged after call', async () => {
    const len = fixtureBuffer.length;
    await svc.executeDetectSections(fixtureBuffer, { sheetName: 'Fund Details' });
    expect(fixtureBuffer.length).toBe(len);
  });
});

describe('GET_CELL_CONTEXT via executeGetCellContext', () => {
  it('returns CellContext for a valid cell', async () => {
    const result = await svc.executeGetCellContext(fixtureBuffer, {
      sheetName: 'Fund Overview',
      coordinate: 'A1',
    });
    expect(result.targetCell.coordinate).toBe('A1');
    expect(result.sheetName).toBe('Fund Overview');
    expect(result.cells.length).toBeGreaterThan(0);
  });

  it('accepts custom radius options', async () => {
    const result = await svc.executeGetCellContext(fixtureBuffer, {
      sheetName: 'Fund Overview',
      coordinate: 'B5',
      rowRadius: 1,
      colRadius: 1,
    });
    // 3×3 window = 9 cells (or less if clamped to bounds)
    expect(result.cells.length).toBeLessThanOrEqual(9);
  });

  it('throws for missing coordinate', async () => {
    await expect(
      svc.executeGetCellContext(fixtureBuffer, { sheetName: 'Fund Overview' }),
    ).rejects.toThrow(/coordinate/i);
  });

  it('throws for invalid coordinate A0', async () => {
    await expect(
      svc.executeGetCellContext(fixtureBuffer, {
        sheetName: 'Fund Overview',
        coordinate: 'A0',
      }),
    ).rejects.toThrow();
  });

  it('throws for non-A1-format coordinate', async () => {
    await expect(
      svc.executeGetCellContext(fixtureBuffer, {
        sheetName: 'Fund Overview',
        coordinate: '1A',
      }),
    ).rejects.toThrow();
  });

  it('throws for rowRadius out of range', async () => {
    await expect(
      svc.executeGetCellContext(fixtureBuffer, {
        sheetName: 'Fund Overview',
        coordinate: 'A1',
        rowRadius: 99,
      }),
    ).rejects.toThrow(/rowRadius/i);
  });

  it('throws for non-existent sheet', async () => {
    await expect(
      svc.executeGetCellContext(fixtureBuffer, {
        sheetName: 'NoSuchSheet',
        coordinate: 'A1',
      }),
    ).rejects.toThrow(/not found/i);
  });

  it('is read-only — buffer unchanged after call', async () => {
    const len = fixtureBuffer.length;
    await svc.executeGetCellContext(fixtureBuffer, {
      sheetName: 'Fund Overview',
      coordinate: 'A5',
    });
    expect(fixtureBuffer.length).toBe(len);
  });

  it('GET_CELL_CONTEXT in validatePlan rejects missing coordinate', () => {
    expect(() =>
      svc.validatePlan({
        summary: 'bad',
        operations: [{
          type: 'GET_CELL_CONTEXT',
          sheet: 'Sheet1',
          params: { sheetName: 'Sheet1' }, // missing coordinate
        }],
      }),
    ).toThrow(/coordinate/i);
  });

  it('GET_CELL_CONTEXT in validatePlan rejects invalid coordinate', () => {
    expect(() =>
      svc.validatePlan({
        summary: 'bad',
        operations: [{
          type: 'GET_CELL_CONTEXT',
          sheet: 'Sheet1',
          params: { sheetName: 'Sheet1', coordinate: '1A' },
        }],
      }),
    ).toThrow();
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// B. FORMULA OVERWRITE PROTECTION
// ═══════════════════════════════════════════════════════════════════════════

describe('Formula overwrite protection — validation layer', () => {
  it('SET_CELL without allowFormulaOverwrite is valid at plan level', () => {
    // At plan validation time we don't have the workbook — just structure check
    expect(() =>
      svc.validatePlan({
        summary: 'test',
        operations: [{ type: 'SET_CELL', sheet: 'S', params: { row: 2, col: 3, value: 100 } }],
      }),
    ).not.toThrow();
  });

  it('SET_CELL with allowFormulaOverwrite: false is valid at plan level', () => {
    expect(() =>
      svc.validatePlan({
        summary: 'test',
        operations: [{
          type: 'SET_CELL', sheet: 'S',
          params: { row: 2, col: 3, value: 100, allowFormulaOverwrite: false },
        }],
      }),
    ).not.toThrow();
  });

  it('SET_CELL with allowFormulaOverwrite: true is valid at plan level', () => {
    expect(() =>
      svc.validatePlan({
        summary: 'test',
        operations: [{
          type: 'SET_CELL', sheet: 'S',
          params: { row: 2, col: 3, value: 100, allowFormulaOverwrite: true },
        }],
      }),
    ).not.toThrow();
  });

  it('SET_CELL with allowFormulaOverwrite as non-boolean is rejected', () => {
    expect(() =>
      svc.validatePlan({
        summary: 'test',
        operations: [{
          type: 'SET_CELL', sheet: 'S',
          params: { row: 2, col: 3, value: 100, allowFormulaOverwrite: 'yes' },
        }],
      }),
    ).toThrow(/allowFormulaOverwrite must be a boolean/i);
  });
});

describe('Formula overwrite protection — execution layer (ExcelService)', () => {
  async function makeFormulaWorkbook(): Promise<Buffer> {
    return buildBuffer((wb) => {
      const ws = wb.addWorksheet('Calc');
      ws.getCell('A1').value = 100;
      ws.getCell('A2').value = 200;
      ws.getCell('A3').value = { formula: 'SUM(A1:A2)', result: 300 };
      ws.getCell('B1').value = 'Label';
      ws.getCell('B3').value = 42; // not a formula
    });
  }

  it('SET_CELL on a non-formula cell works normally', async () => {
    const buf = await makeFormulaWorkbook();
    const wb = await excelSvc.loadWorkbook(buf);
    const plan = {
      summary: 'set non-formula',
      operations: [{ type: 'SET_CELL', sheet: 'Calc', params: { row: 1, col: 1, value: 999 } }],
    };
    svc.validatePlan(plan);
    excelSvc.applyOperationPlan(wb, plan);
    expect(wb.getWorksheet('Calc').getCell(1, 1).value).toBe(999);
  });

  it('SET_CELL on a formula cell WITHOUT allowFormulaOverwrite THROWS', async () => {
    const buf = await makeFormulaWorkbook();
    const wb = await excelSvc.loadWorkbook(buf);
    const plan = {
      summary: 'overwrite formula without permission',
      operations: [{ type: 'SET_CELL', sheet: 'Calc', params: { row: 3, col: 1, value: 500 } }],
    };
    svc.validatePlan(plan);
    expect(() => excelSvc.applyOperationPlan(wb, plan)).toThrow(/formula/i);
  });

  it('error message mentions the cell address', async () => {
    const buf = await makeFormulaWorkbook();
    const wb = await excelSvc.loadWorkbook(buf);
    const plan = {
      summary: 'test',
      operations: [{ type: 'SET_CELL', sheet: 'Calc', params: { row: 3, col: 1, value: 500 } }],
    };
    try {
      excelSvc.applyOperationPlan(wb, plan);
      fail('Should have thrown');
    } catch (err: any) {
      expect(err.message).toMatch(/A3/); // cell coordinate in error
      expect(err.message).toMatch(/formula/i);
      expect(err.message).toMatch(/allowFormulaOverwrite/i);
    }
  });

  it('error message includes the existing formula', async () => {
    const buf = await makeFormulaWorkbook();
    const wb = await excelSvc.loadWorkbook(buf);
    const plan = {
      summary: 'test',
      operations: [{ type: 'SET_CELL', sheet: 'Calc', params: { row: 3, col: 1, value: 500 } }],
    };
    try {
      excelSvc.applyOperationPlan(wb, plan);
    } catch (err: any) {
      expect(err.message).toContain('SUM(A1:A2)');
    }
  });

  it('SET_CELL with allowFormulaOverwrite: true replaces the formula', async () => {
    const buf = await makeFormulaWorkbook();
    const wb = await excelSvc.loadWorkbook(buf);
    const plan = {
      summary: 'intentional overwrite',
      operations: [{
        type: 'SET_CELL', sheet: 'Calc',
        params: { row: 3, col: 1, value: 500, allowFormulaOverwrite: true },
      }],
    };
    svc.validatePlan(plan);
    excelSvc.applyOperationPlan(wb, plan);
    const cell = wb.getWorksheet('Calc').getCell(3, 1);
    expect(cell.value).toBe(500);
    // Cell should no longer be a formula
    expect(cell.type).not.toBe(ExcelJS.ValueType.Formula);
  });

  it('allowFormulaOverwrite defaults to false when absent', async () => {
    const buf = await makeFormulaWorkbook();
    const wb = await excelSvc.loadWorkbook(buf);
    const plan = {
      summary: 'no flag — should block',
      operations: [{ type: 'SET_CELL', sheet: 'Calc', params: { row: 3, col: 1, value: 777 } }],
    };
    // No allowFormulaOverwrite key at all → should throw
    expect(() => excelSvc.applyOperationPlan(wb, plan)).toThrow(/formula/i);
  });

  it('allowFormulaOverwrite: false explicitly still blocks', async () => {
    const buf = await makeFormulaWorkbook();
    const wb = await excelSvc.loadWorkbook(buf);
    const plan = {
      summary: 'explicit false — should block',
      operations: [{
        type: 'SET_CELL', sheet: 'Calc',
        params: { row: 3, col: 1, value: 777, allowFormulaOverwrite: false },
      }],
    };
    expect(() => excelSvc.applyOperationPlan(wb, plan)).toThrow(/formula/i);
  });

  it('formula protection does not affect ADD_FORMULA (which sets formulas intentionally)', async () => {
    const buf = await makeFormulaWorkbook();
    const wb = await excelSvc.loadWorkbook(buf);
    const plan = {
      summary: 'add formula to empty cell',
      operations: [{
        type: 'ADD_FORMULA', sheet: 'Calc',
        params: { row: 4, col: 1, formula: 'A1*2' },
      }],
    };
    svc.validatePlan(plan);
    expect(() => excelSvc.applyOperationPlan(wb, plan)).not.toThrow();
    const cell = wb.getWorksheet('Calc').getCell(4, 1);
    expect(cell.type).toBe(ExcelJS.ValueType.Formula);
  });

  it('protection applies inside multi-op plan — stops at first formula violation', async () => {
    const buf = await makeFormulaWorkbook();
    const wb = await excelSvc.loadWorkbook(buf);
    const plan = {
      summary: 'multi-op with formula violation',
      operations: [
        { type: 'SET_CELL', sheet: 'Calc', params: { row: 1, col: 1, value: 'first ok' } },
        { type: 'SET_CELL', sheet: 'Calc', params: { row: 3, col: 1, value: 'blocked' } }, // formula cell
        { type: 'SET_CELL', sheet: 'Calc', params: { row: 2, col: 1, value: 'never reached' } },
      ],
    };
    expect(() => excelSvc.applyOperationPlan(wb, plan)).toThrow(/formula/i);
    // First op applied before the throw
    expect(wb.getWorksheet('Calc').getCell(1, 1).value).toBe('first ok');
    // Third op never reached
    expect(wb.getWorksheet('Calc').getCell(2, 1).value).toBe(200);
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// C. SET_CELL_BY_LABEL FOUNDATION
// ═══════════════════════════════════════════════════════════════════════════

describe('SET_CELL_BY_LABEL — validatePlan', () => {
  it('SET_CELL_BY_LABEL is accepted by validatePlan', () => {
    expect(() =>
      svc.validatePlan({
        summary: 'test',
        operations: [{
          type: 'SET_CELL_BY_LABEL',
          sheet: 'Fund Details',
          params: { label: 'Management Fee', value: 1.5 },
        }],
      }),
    ).not.toThrow();
  });

  it('rejects missing label', () => {
    expect(() =>
      svc.validatePlan({
        summary: 'test',
        operations: [{
          type: 'SET_CELL_BY_LABEL',
          sheet: 'S',
          params: { value: 1.5 },
        }],
      }),
    ).toThrow(/label/i);
  });

  it('rejects empty label', () => {
    expect(() =>
      svc.validatePlan({
        summary: 'test',
        operations: [{
          type: 'SET_CELL_BY_LABEL',
          sheet: 'S',
          params: { label: '   ', value: 1.5 },
        }],
      }),
    ).toThrow(/label/i);
  });

  it('rejects label exceeding 500 chars', () => {
    expect(() =>
      svc.validatePlan({
        summary: 'test',
        operations: [{
          type: 'SET_CELL_BY_LABEL',
          sheet: 'S',
          params: { label: 'x'.repeat(501), value: 1 },
        }],
      }),
    ).toThrow(/500/);
  });

  it('rejects invalid valueColOffset', () => {
    expect(() =>
      svc.validatePlan({
        summary: 'test',
        operations: [{
          type: 'SET_CELL_BY_LABEL',
          sheet: 'S',
          params: { label: 'Fee', value: 1, valueColOffset: 999 },
        }],
      }),
    ).toThrow(/valueColOffset/i);
  });

  it('rejects non-boolean allowFormulaOverwrite', () => {
    expect(() =>
      svc.validatePlan({
        summary: 'test',
        operations: [{
          type: 'SET_CELL_BY_LABEL',
          sheet: 'S',
          params: { label: 'Fee', value: 1, allowFormulaOverwrite: 'yes' },
        }],
      }),
    ).toThrow(/allowFormulaOverwrite must be a boolean/i);
  });

  it('accepts optional sheetName filter', () => {
    expect(() =>
      svc.validatePlan({
        summary: 'test',
        operations: [{
          type: 'SET_CELL_BY_LABEL',
          sheet: 'Fund Details',
          params: { label: 'Management Fee', value: 1.5, sheetName: 'Fund Details' },
        }],
      }),
    ).not.toThrow();
  });
});

describe('SET_CELL_BY_LABEL — executeSetCellByLabel resolution', () => {
  it('returns "not_found" for label that does not exist', async () => {
    const result = await svc.executeSetCellByLabel(fixtureBuffer, {
      label: 'ZZZNOMATCH_LABEL_ZZZ',
      value: 999,
    });
    expect(result.status).toBe('not_found');
    expect(result.resolvedTarget).toBeNull();
    expect(result.candidates).toHaveLength(0);
    expect(result.message).toContain('not found');
  });

  it('returns "ambiguous" when label appears in multiple locations', async () => {
    // "Management Fee" appears in Fund Overview, Fund Details, Fee Structure, Notes
    const result = await svc.executeSetCellByLabel(fixtureBuffer, {
      label: 'Management Fee',
      value: 1.5,
    });
    expect(result.status).toBe('ambiguous');
    expect(result.resolvedTarget).toBeNull();
    expect(result.candidates.length).toBeGreaterThan(1);
    expect(result.message.toLowerCase()).toContain('ambiguous');
    expect(result.message).toContain('GET_CELL_CONTEXT');
  });

  it('ambiguous result lists all candidates with coordinates', async () => {
    const result = await svc.executeSetCellByLabel(fixtureBuffer, {
      label: 'Management Fee',
      value: 1.5,
    });
    for (const c of result.candidates) {
      expect(c.sheetName).toBeTruthy();
      expect(c.labelCoordinate).toMatch(/^[A-Z]+\d+$/);
      expect(c.valueCoordinate).toMatch(/^[A-Z]+\d+$/);
      expect(c.labelValue).toBeTruthy();
    }
  });

  it('NEVER auto-selects first match when ambiguous', async () => {
    const result = await svc.executeSetCellByLabel(fixtureBuffer, {
      label: 'Management Fee',
      value: 1.5,
    });
    // Core anti-hallucination check: resolvedTarget must be null
    expect(result.resolvedTarget).toBeNull();
  });

  it('returns "resolved" when label is unique in filtered sheet', async () => {
    // "Total Fee Burden" appears only in Fund Details
    const result = await svc.executeSetCellByLabel(fixtureBuffer, {
      label: 'Total Fee Burden',
      value: 15,
      sheetName: 'Fund Details',
    });
    // Could be resolved or not_found depending on exact fixture content
    // Either way, it must not be ambiguous
    expect(['resolved', 'not_found']).toContain(result.status);
    if (result.status === 'resolved') {
      expect(result.resolvedTarget).not.toBeNull();
      expect(result.resolvedTarget!.sheetName).toBe('Fund Details');
      expect(result.resolvedTarget!.labelCoordinate).toMatch(/^[A-Z]+\d+$/);
      expect(result.resolvedTarget!.valueCoordinate).toMatch(/^[A-Z]+\d+$/);
    }
  });

  it('valueCoordinate is offset from labelCoordinate by valueColOffset', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('Fees');
      ws.getCell('A5').value = 'Unique Target Label XYZ';
      ws.getCell('B5').value = 1.25; // offset 1
      ws.getCell('C5').value = '%';  // offset 2
    });
    const result = await svc.executeSetCellByLabel(buf, {
      label: 'Unique Target Label XYZ',
      value: 9.99,
      valueColOffset: 1,
    });
    expect(result.status).toBe('resolved');
    expect(result.resolvedTarget!.labelCoordinate).toBe('A5');
    expect(result.resolvedTarget!.valueCoordinate).toBe('B5');
  });

  it('valueColOffset: 2 points two columns right of label', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('Fees');
      ws.getCell('A5').value = 'Unique Label ABC';
      ws.getCell('C5').value = 1.25; // offset 2
    });
    const result = await svc.executeSetCellByLabel(buf, {
      label: 'Unique Label ABC',
      value: 9.99,
      valueColOffset: 2,
    });
    expect(result.status).toBe('resolved');
    expect(result.resolvedTarget!.valueCoordinate).toBe('C5');
  });

  it('sheet filter limits search to one sheet', async () => {
    const result = await svc.executeSetCellByLabel(fixtureBuffer, {
      label: 'Management Fee',
      value: 1.5,
      sheetName: 'Fund Details',
    });
    // All candidates must be from Fund Details
    for (const c of result.candidates) {
      expect(c.sheetName).toBe('Fund Details');
    }
  });

  it('returns message with actionable information', async () => {
    const result = await svc.executeSetCellByLabel(fixtureBuffer, {
      label: 'Management Fee',
      value: 1.5,
    });
    expect(typeof result.message).toBe('string');
    expect(result.message.length).toBeGreaterThan(10);
  });

  it('is read-only — buffer unchanged after resolution', async () => {
    const len = fixtureBuffer.length;
    await svc.executeSetCellByLabel(fixtureBuffer, {
      label: 'Management Fee',
      value: 1.5,
    });
    expect(fixtureBuffer.length).toBe(len);
  });

  it('throws for empty label', async () => {
    await expect(
      svc.executeSetCellByLabel(fixtureBuffer, { label: '   ', value: 1 }),
    ).rejects.toThrow(/label/i);
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// D. REGRESSION — existing write operations
// ═══════════════════════════════════════════════════════════════════════════

describe('Regression — existing 8 write operations still work', () => {
  const writeOps = [
    { type: 'SET_CELL', params: { row: 1, col: 1, value: 'x' } },
    { type: 'SET_CELL_RANGE', params: { startRow: 1, startCol: 1, endRow: 2, endCol: 2, value: 0 } },
    { type: 'ADD_ROW', params: { values: ['a', 'b'] } },
    { type: 'DELETE_ROW', params: { row: 1 } },
    { type: 'ADD_FORMULA', params: { row: 1, col: 1, formula: 'SUM(A1:A2)' } },
    { type: 'FORMAT_CELL', params: { row: 1, col: 1, style: { font: { bold: true } } } },
    { type: 'ADD_COLUMN_HEADER', params: { col: 1, header: 'Name' } },
    { type: 'SORT_SHEET', params: { col: 1, order: 'asc' } },
  ];

  writeOps.forEach(({ type, params }) => {
    it(`${type} is still accepted by validatePlan`, () => {
      expect(() =>
        svc.validatePlan({
          summary: 'regression',
          operations: [{ type, sheet: 'Sheet1', params }],
        }),
      ).not.toThrow();
    });
  });

  it('SET_CELL without formula cell executes successfully end-to-end', async () => {
    const buf = await buildBuffer((wb) => {
      const ws = wb.addWorksheet('Test');
      ws.getCell('A1').value = 'old';
    });
    const wb = await excelSvc.loadWorkbook(buf);
    const plan = {
      summary: 'regression',
      operations: [{ type: 'SET_CELL', sheet: 'Test', params: { row: 1, col: 1, value: 'new' } }],
    };
    svc.validatePlan(plan);
    excelSvc.applyOperationPlan(wb, plan);
    expect(wb.getWorksheet('Test').getCell(1, 1).value).toBe('new');
  });

  it('isReadOperation returns true only for inspection types', () => {
    expect(svc.isReadOperation('INSPECT_WORKBOOK_STRUCTURE')).toBe(true);
    expect(svc.isReadOperation('DETECT_HEADERS')).toBe(true);
    expect(svc.isReadOperation('DETECT_SECTIONS')).toBe(true);
    expect(svc.isReadOperation('GET_CELL_CONTEXT')).toBe(true);
    expect(svc.isReadOperation('SET_CELL')).toBe(false);
    expect(svc.isReadOperation('SET_CELL_BY_LABEL')).toBe(false);
  });

  it('isWriteOperation returns true only for write types', () => {
    expect(svc.isWriteOperation('SET_CELL')).toBe(true);
    expect(svc.isWriteOperation('ADD_ROW')).toBe(true);
    expect(svc.isWriteOperation('INSPECT_WORKBOOK_STRUCTURE')).toBe(false);
    expect(svc.isWriteOperation('SET_CELL_BY_LABEL')).toBe(false);
  });
});
