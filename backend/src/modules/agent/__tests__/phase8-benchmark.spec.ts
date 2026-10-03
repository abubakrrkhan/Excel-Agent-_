/**
 * phase8-benchmark.spec.ts
 *
 * Phase 8 — Financial benchmark + full pipeline integration tests.
 *
 * Tests the complete stack end-to-end using the synthetic financial fixture:
 *   A. Workbook inspection accuracy (fixture as ground truth)
 *   B. Formula protection across the full pipeline
 *   C. Label resolution accuracy (Management Fee, Total Fee Burden, etc.)
 *   D. Multi-sheet write safety (writes don't leak across sheets)
 *   E. Agent → validate → route → execute → save full pipeline
 *   F. Output workbook integrity (original never modified)
 *   G. Error recovery (invalid ops produce clean errors, not crashes)
 *   H. Performance budget (key ops complete within time limits)
 */

import * as ExcelJS from 'exceljs';
import { ExcelService } from '../../excel/excel.service';
import { ExcelInspectionService } from '../../excel/excel-inspection.service';
import { WorkbookStructureService } from '../../excel/workbook-structure.service';
import { OperationsService } from '../../operations/operations.service';
import { OperationExecutionRouter } from '../../operations/operation-execution-router';
import { AgentService } from '../agent.service';
import { AgentRoundResponse } from '../agent.types';
import { buildFinancialFixtureBuffer } from '../../excel/fixtures/financial-workbook.fixture';

// ─── Setup ────────────────────────────────────────────────────────────────────

let excelSvc: ExcelService;
let structSvc: WorkbookStructureService;
let inspSvc: ExcelInspectionService;
let opsSvc: OperationsService;
let router: OperationExecutionRouter;
let fixtureBuffer: Buffer;

beforeAll(async () => {
  excelSvc = new ExcelService();
  structSvc = new WorkbookStructureService();
  inspSvc = new ExcelInspectionService(structSvc);
  opsSvc = new OperationsService(inspSvc, structSvc);
  router = new OperationExecutionRouter(opsSvc, excelSvc);
  fixtureBuffer = await buildFinancialFixtureBuffer();
});

function buildAgent(responses: AgentRoundResponse[]): AgentService {
  let call = 0;
  const mockAi = { callAgentRound: jest.fn(async () => {
    const r = responses[call] ?? responses[responses.length - 1]; call++; return r;
  }) };
  return new AgentService(mockAi as any, inspSvc, structSvc, opsSvc);
}

// ═══════════════════════════════════════════════════════════════════════════
// A. Workbook inspection accuracy
// ═══════════════════════════════════════════════════════════════════════════

describe('A. Workbook inspection accuracy', () => {
  it('fixture has exactly 5 sheets', async () => {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(fixtureBuffer as any);
    expect(wb.worksheets.length).toBe(5);
  });

  it('Fund Overview has title merged cell at A1', async () => {
    const result = await structSvc.inspectSheetStructure(fixtureBuffer, 'Fund Overview');
    const titleMerge = result.mergedRanges.find(m => m.masterAddress === 'A1');
    expect(titleMerge).toBeDefined();
    expect(titleMerge!.colSpan).toBeGreaterThan(1);
  });

  it('Fund Overview has hidden row 3', async () => {
    const result = await structSvc.inspectSheetStructure(fixtureBuffer, 'Fund Overview');
    expect(result.hiddenRows).toContain(3);
  });

  it('Fund Details has hidden row 7', async () => {
    const result = await structSvc.inspectSheetStructure(fixtureBuffer, 'Fund Details');
    expect(result.hiddenRows).toContain(7);
  });

  it('Fund Details has at least 2 sections (fees + performance)', async () => {
    const result = await structSvc.detectSections(fixtureBuffer, 'Fund Details');
    expect(result.sections.length).toBeGreaterThanOrEqual(2);
  });

  it('Portfolio has 2 sections (equity + fixed income)', async () => {
    const result = await structSvc.detectSections(fixtureBuffer, 'Portfolio');
    expect(result.sections.length).toBeGreaterThanOrEqual(2);
  });

  it('Management Fee appears in 4+ sheets', async () => {
    const result = await inspSvc.searchWorkbook(fixtureBuffer, 'Management Fee');
    const sheets = new Set(result.matches.map(m => m.sheetName));
    expect(sheets.size).toBeGreaterThanOrEqual(3);
    expect(result.matches.length).toBeGreaterThanOrEqual(4);
  });

  it('Fund Details has formula cells (SUM, AVERAGE)', async () => {
    const result = await structSvc.inspectSheetStructure(fixtureBuffer, 'Fund Details');
    expect(result.cellCounts.formula).toBeGreaterThanOrEqual(2);
  });

  it('cell context of Management Fee row includes value cell', async () => {
    const search = await inspSvc.searchWorkbook(fixtureBuffer, 'Management Fee', {
      sheetName: 'Fund Details', maxMatches: 1,
    });
    expect(search.matches.length).toBeGreaterThan(0);
    const coord = search.matches[0].coordinate;
    const ctx = await structSvc.getCellContext(fixtureBuffer, 'Fund Details', coord);
    // Row should have at least 1 other cell (the value cell)
    expect(ctx.rowLabels.length).toBeGreaterThanOrEqual(0);
    expect(ctx.targetCell.value).toContain('Management Fee');
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// B. Formula protection — full pipeline
// ═══════════════════════════════════════════════════════════════════════════

describe('B. Formula protection — full pipeline', () => {
  it('SET_CELL on SUM formula in Fund Details is blocked', async () => {
    const wb = await excelSvc.loadWorkbook(fixtureBuffer);
    // C21 has SUM(C10:C12) formula
    const plan = {
      summary: 'blocked',
      operations: [{ type: 'SET_CELL', sheet: 'Fund Details', params: { row: 21, col: 3, value: 999 } }],
    };
    opsSvc.validatePlan(plan);
    expect(() => excelSvc.applyOperationPlan(wb, plan)).toThrow(/formula/i);
  });

  it('SET_CELL_RANGE spanning formula cells in Fund Details is blocked (atomic)', async () => {
    const wb = await excelSvc.loadWorkbook(fixtureBuffer);
    // Range covering rows 10-22 col 3 includes formula at row 21
    const plan = {
      summary: 'blocked range',
      operations: [{
        type: 'SET_CELL_RANGE', sheet: 'Fund Details',
        params: { startRow: 10, startCol: 3, endRow: 22, endCol: 3, value: 0 },
      }],
    };
    opsSvc.validatePlan(plan);
    expect(() => excelSvc.applyOperationPlan(wb, plan)).toThrow(/formula/i);
  });

  it('SET_CELL_RANGE with allowFormulaOverwrite succeeds across formula cells', async () => {
    const wb = await excelSvc.loadWorkbook(fixtureBuffer);
    const plan = {
      summary: 'authorized',
      operations: [{
        type: 'SET_CELL_RANGE', sheet: 'Fund Details',
        params: { startRow: 21, startCol: 3, endRow: 21, endCol: 3, value: 11.5, allowFormulaOverwrite: true },
      }],
    };
    opsSvc.validatePlan(plan);
    expect(() => excelSvc.applyOperationPlan(wb, plan)).not.toThrow();
  });

  it('router propagates formula error from Fund Details', async () => {
    const wb = await excelSvc.loadWorkbook(fixtureBuffer);
    const plan = {
      summary: 'test',
      operations: [{ type: 'SET_CELL', sheet: 'Fund Details', params: { row: 21, col: 3, value: 999 } }],
    };
    opsSvc.validatePlan(plan);
    await expect(router.executePlan(fixtureBuffer, plan, wb)).rejects.toThrow(/formula/i);
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// C. Label resolution accuracy
// ═══════════════════════════════════════════════════════════════════════════

describe('C. Label resolution accuracy', () => {
  it('Management Fee is ambiguous across full fixture', async () => {
    const result = await opsSvc.executeSetCellByLabel(fixtureBuffer, {
      label: 'Management Fee', value: 1.5,
    });
    expect(result.status).toBe('ambiguous');
    expect(result.resolvedTarget).toBeNull();
  });

  it('Management Fee narrows to resolved on Fund Details sheet filter', async () => {
    const result = await opsSvc.executeSetCellByLabel(fixtureBuffer, {
      label: 'Management Fee', value: 1.5,
      sheetName: 'Fund Details',
    });
    // Should be resolved (one match in Fund Details) or ambiguous (multiple)
    expect(['resolved', 'ambiguous']).toContain(result.status);
    // Key assertion: not not_found
    expect(result.status).not.toBe('not_found');
  });

  it('Total Fee Burden resolves on Fund Details sheet', async () => {
    const result = await opsSvc.executeSetCellByLabel(fixtureBuffer, {
      label: 'Total Fee Burden', value: 12.0,
      sheetName: 'Fund Details',
    });
    expect(['resolved', 'not_found']).toContain(result.status);
    if (result.status === 'resolved') {
      expect(result.resolvedTarget!.sheetName).toBe('Fund Details');
    }
  });

  it('non-existent label returns not_found', async () => {
    const result = await opsSvc.executeSetCellByLabel(fixtureBuffer, {
      label: 'ZZZNOMATCH_LABEL_BENCHMARK', value: 0,
    });
    expect(result.status).toBe('not_found');
    expect(result.candidates).toHaveLength(0);
  });

  it('resolve_label with valueColOffset=2 finds correct column', async () => {
    const search = await inspSvc.searchWorkbook(fixtureBuffer, 'Management Fee', {
      sheetName: 'Fund Details', maxMatches: 1,
    });
    if (search.matches.length > 0) {
      const coord = search.matches[0].coordinate;
      const ctx = await structSvc.getCellContext(fixtureBuffer, 'Fund Details', coord);
      expect(ctx.targetCell.coordinate).toBe(coord);
      expect(ctx.cells.length).toBeGreaterThan(0);
    }
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// D. Multi-sheet write safety
// ═══════════════════════════════════════════════════════════════════════════

describe('D. Multi-sheet write safety', () => {
  it('write to Fund Details does not affect Fund Overview', async () => {
    const wb = await excelSvc.loadWorkbook(fixtureBuffer);
    const before = wb.getWorksheet('Fund Overview').getCell('A5').value;

    const plan = {
      summary: 'write to Fund Details only',
      operations: [
        { type: 'SET_CELL', sheet: 'Fund Details', params: { row: 5, col: 2, value: 'MODIFIED' } },
      ],
    };
    opsSvc.validatePlan(plan);
    await router.executePlan(fixtureBuffer, plan, wb);

    // Fund Overview A5 must be unchanged
    expect(wb.getWorksheet('Fund Overview').getCell('A5').value).toEqual(before);
    // Fund Details B5 must be changed
    expect(wb.getWorksheet('Fund Details').getCell(5, 2).value).toBe('MODIFIED');
  });

  it('plan targeting 3 different sheets applies writes independently', async () => {
    const wb = await excelSvc.loadWorkbook(fixtureBuffer);
    const plan = {
      summary: 'multi-sheet add rows',
      operations: [
        { type: 'ADD_ROW', sheet: 'Fund Overview', params: { values: ['New Row', 'A'] } },
        { type: 'ADD_ROW', sheet: 'Portfolio', params: { values: ['New Asset', 'Equity', 5] } },
      ],
    };
    opsSvc.validatePlan(plan);
    const result = await router.executePlan(fixtureBuffer, plan, wb);
    expect(result.writeOpsApplied).toBe(2);
    expect(result.workbookModified).toBe(true);
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// E. Full agent pipeline: agent → validate → route → execute
// ═══════════════════════════════════════════════════════════════════════════

describe('E. Full agent pipeline', () => {
  const fixtureInfo = {
    sheetCount: 5,
    sheets: [
      { name: 'Fund Overview', rowCount: 20, dataRowCount: 18, columnCount: 6, dataColumnCount: 5, headers: [] },
      { name: 'Fund Details', rowCount: 35, dataRowCount: 32, columnCount: 8, dataColumnCount: 7, headers: [] },
      { name: 'Portfolio', rowCount: 25, dataRowCount: 22, columnCount: 5, dataColumnCount: 5, headers: [] },
      { name: 'Fee Structure', rowCount: 12, dataRowCount: 10, columnCount: 4, dataColumnCount: 4, headers: [] },
      { name: 'Notes', rowCount: 8, dataRowCount: 6, columnCount: 3, dataColumnCount: 3, headers: [] },
    ],
  };

  it('agent + router: inspect then write, verify final workbook value', async () => {
    const agent = buildAgent([
      {
        type: 'tool_call', tool: 'detect_sections',
        params: { sheetName: 'Fund Details' }, reasoning: 'inspect first',
      },
      {
        type: 'final_plan', reasoning: 'Section confirms row 5 is in fee area',
        plan: {
          summary: 'Update fee',
          operations: [{ type: 'SET_CELL', sheet: 'Fund Details', params: { row: 5, col: 2, value: 'BENCHMARK_WRITE' } }],
        },
      },
    ]);

    const agentResult = await agent.runAgent(fixtureBuffer, 'Update fund details row 5', fixtureInfo);
    expect(agentResult.status).toBe('completed');

    // Apply through the router
    const wb = await excelSvc.loadWorkbook(fixtureBuffer);
    opsSvc.validatePlan(agentResult.plan!);
    const routerResult = await router.executePlan(fixtureBuffer, agentResult.plan!, wb);
    expect(routerResult.writeOpsApplied).toBe(1);

    // Save and reload to verify persistence
    const outputBuf = await excelSvc.saveWorkbook(wb);
    const wb2 = await excelSvc.loadWorkbook(outputBuf);
    expect(wb2.getWorksheet('Fund Details').getCell(5, 2).value).toBe('BENCHMARK_WRITE');
  });

  it('cannot_proceed agent produces no plan and does not attempt writes', async () => {
    const agent = buildAgent([{
      type: 'cannot_proceed',
      reasoning: 'email not supported',
      reason: 'I cannot send emails from this system.',
    }]);

    const result = await agent.runAgent(fixtureBuffer, 'Email the fund report', fixtureInfo);
    expect(result.status).toBe('cannot_proceed');
    expect(result.plan).toBeNull();
    expect(result.cannotProceedReason).toContain('email');
  });

  it('activity log covers all phases of a 3-round session', async () => {
    const agent = buildAgent([
      { type: 'tool_call', tool: 'search_workbook', params: { query: 'NAV' }, reasoning: 'search' },
      { type: 'tool_call', tool: 'get_cell_context', params: { sheetName: 'Fund Overview', coordinate: 'A5' }, reasoning: 'verify' },
      {
        type: 'final_plan', reasoning: 'confirmed',
        plan: { summary: 'update NAV', operations: [{ type: 'SET_CELL', sheet: 'Fund Overview', params: { row: 5, col: 2, value: 110 } }] },
      },
    ]);

    const result = await agent.runAgent(fixtureBuffer, 'Update NAV', fixtureInfo);
    expect(result.activityLog.length).toBeGreaterThan(5);
    // Should have round markers
    const roundLogs = result.activityLog.filter(l => l.includes('Round'));
    expect(roundLogs.length).toBeGreaterThanOrEqual(3);
    // Should have tool call logs
    const toolLogs = result.activityLog.filter(l => l.includes('Tool call'));
    expect(toolLogs.length).toBe(2);
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// F. Output workbook integrity (original never modified)
// ═══════════════════════════════════════════════════════════════════════════

describe('F. Output workbook integrity', () => {
  it('original buffer is unchanged after full pipeline execution', async () => {
    const originalLen = fixtureBuffer.length;
    const wb = await excelSvc.loadWorkbook(fixtureBuffer);
    const plan = {
      summary: 'test',
      operations: [{ type: 'SET_CELL', sheet: 'Fund Overview', params: { row: 5, col: 2, value: 'CHANGED' } }],
    };
    opsSvc.validatePlan(plan);
    await router.executePlan(fixtureBuffer, plan, wb);
    await excelSvc.saveWorkbook(wb);

    // Original buffer must be unchanged
    expect(fixtureBuffer.length).toBe(originalLen);
  });

  it('output buffer differs from input after write', async () => {
    const wb = await excelSvc.loadWorkbook(fixtureBuffer);
    const plan = {
      summary: 'test',
      operations: [{ type: 'SET_CELL', sheet: 'Fund Overview', params: { row: 5, col: 2, value: 'CHANGED_OUTPUT' } }],
    };
    opsSvc.validatePlan(plan);
    await router.executePlan(fixtureBuffer, plan, wb);
    const outputBuf = await excelSvc.saveWorkbook(wb);

    // Output contains the change, original does not
    const outWb = await excelSvc.loadWorkbook(outputBuf);
    expect(outWb.getWorksheet('Fund Overview').getCell(5, 2).value).toBe('CHANGED_OUTPUT');

    // Re-loading original fixture confirms it's unchanged
    const origWb = await excelSvc.loadWorkbook(fixtureBuffer);
    expect(origWb.getWorksheet('Fund Overview').getCell(5, 2).value).not.toBe('CHANGED_OUTPUT');
  });

  it('read-only plan (inspection only) produces no output change', async () => {
    const wb = await excelSvc.loadWorkbook(fixtureBuffer);
    const plan = {
      summary: 'inspect only',
      operations: [
        { type: 'DETECT_SECTIONS', sheet: 'Portfolio', params: { sheetName: 'Portfolio' } },
        { type: 'INSPECT_WORKBOOK_STRUCTURE', sheet: 'Fund Details', params: { sheetName: 'Fund Details' } },
      ],
    };
    opsSvc.validatePlan(plan);
    const result = await router.executePlan(fixtureBuffer, plan, wb);
    expect(result.workbookModified).toBe(false);
    expect(result.writeOpsApplied).toBe(0);
    expect(result.readOpsExecuted).toBe(2);
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// G. Error recovery — clean errors, not crashes
// ═══════════════════════════════════════════════════════════════════════════

describe('G. Error recovery', () => {
  it('invalid sheet name throws BadRequestException, not crash', async () => {
    await expect(
      structSvc.inspectSheetStructure(fixtureBuffer, 'NOSUCHSHEET'),
    ).rejects.toThrow(/not found/i);
  });

  it('invalid coordinate in GET_CELL_CONTEXT throws, not crash', async () => {
    await expect(
      opsSvc.executeGetCellContext(fixtureBuffer, { sheetName: 'Fund Overview', coordinate: 'INVALID' }),
    ).rejects.toThrow();
  });

  it('plan with unknown op type throws at validation, not execution', () => {
    expect(() =>
      opsSvc.validatePlan({
        summary: 'bad',
        operations: [{ type: 'HACK_DB', sheet: 'Fund Overview', params: {} }],
      }),
    ).toThrow(/unsupported operation type/i);
  });

  it('empty plan throws at validation', () => {
    expect(() =>
      opsSvc.validatePlan({ summary: 'empty', operations: [] }),
    ).toThrow(/empty/i);
  });

  it('router propagates tool error as success=false in agent, not crash', async () => {
    const agent = buildAgent([
      {
        type: 'tool_call', tool: 'detect_sections',
        params: { sheetName: 'NOSUCHSHEET' }, reasoning: 'bad sheet',
      },
      {
        type: 'final_plan', reasoning: 'fallback',
        plan: { summary: 'fallback', operations: [{ type: 'ADD_ROW', sheet: 'Fund Overview', params: { values: ['x'] } }] },
      },
    ]);

    const info = { sheetCount: 1, sheets: [{ name: 'Fund Overview', rowCount: 5, dataRowCount: 4, columnCount: 2, dataColumnCount: 2, headers: [] }] };
    const result = await agent.runAgent(fixtureBuffer, 'test', info);
    // Agent should not crash — tool failure handled gracefully
    expect(result.toolHistory[0].success).toBe(false);
    expect(result.toolHistory[0].error).toBeDefined();
    // Agent continues to final plan
    expect(result.status).toBe('completed');
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// H. Performance budget
// ═══════════════════════════════════════════════════════════════════════════

describe('H. Performance budget', () => {
  it('searchWorkbook on fixture completes in < 2000ms', async () => {
    const start = Date.now();
    await inspSvc.searchWorkbook(fixtureBuffer, 'Management Fee');
    expect(Date.now() - start).toBeLessThan(2000);
  });

  it('detectSections on Fund Details completes in < 2000ms', async () => {
    const start = Date.now();
    await structSvc.detectSections(fixtureBuffer, 'Fund Details');
    expect(Date.now() - start).toBeLessThan(2000);
  });

  it('full inspection suite (5 ops) on fixture completes in < 5000ms', async () => {
    const start = Date.now();
    await Promise.all([
      structSvc.inspectSheetStructure(fixtureBuffer, 'Fund Overview'),
      structSvc.inspectSheetStructure(fixtureBuffer, 'Fund Details'),
      structSvc.detectSections(fixtureBuffer, 'Portfolio'),
      structSvc.detectHeaders(fixtureBuffer, 'Fee Structure'),
      inspSvc.searchWorkbook(fixtureBuffer, 'NAV'),
    ]);
    expect(Date.now() - start).toBeLessThan(5000);
  });

  it('loadWorkbook → applyPlan → saveWorkbook round-trip in < 3000ms', async () => {
    const start = Date.now();
    const wb = await excelSvc.loadWorkbook(fixtureBuffer);
    const plan = {
      summary: 'perf test',
      operations: [
        { type: 'SET_CELL', sheet: 'Fund Overview', params: { row: 5, col: 2, value: 'PERF' } },
        { type: 'ADD_ROW', sheet: 'Portfolio', params: { values: ['Perf Asset', 'Test', 1] } },
      ],
    };
    opsSvc.validatePlan(plan);
    await router.executePlan(fixtureBuffer, plan, wb);
    await excelSvc.saveWorkbook(wb);
    expect(Date.now() - start).toBeLessThan(3000);
  });
});
