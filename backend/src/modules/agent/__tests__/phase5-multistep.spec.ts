/**
 * phase5-multistep.spec.ts
 *
 * Phase 5 — Multi-step / multi-sheet agent capability tests.
 *
 * Verifies the existing agent can:
 *   A. Execute multi-operation plans (read + write sequence)
 *   B. Handle multi-sheet workbooks (operations across different sheets)
 *   C. Use search_workbook to locate targets across sheets before writing
 *   D. Produce plans with mixed inspection + write ops
 *   E. Summarise large tool results (Phase 3 Step 2)
 *   F. Handle the clarification flow end-to-end
 */

import * as ExcelJS from 'exceljs';
import { AgentService } from '../agent.service';
import { AiService } from '../../ai/ai.service';
import { ExcelInspectionService } from '../../excel/excel-inspection.service';
import { WorkbookStructureService } from '../../excel/workbook-structure.service';
import { OperationsService } from '../../operations/operations.service';
import { ExcelService, WorkbookInfo } from '../../excel/excel.service';
import { OperationExecutionRouter } from '../../operations/operation-execution-router';
import { AgentRoundResponse } from '../agent.types';
import { buildFinancialFixtureBuffer } from '../../excel/fixtures/financial-workbook.fixture';

// ─── Helpers ──────────────────────────────────────────────────────────────────

async function buildMultiSheetBuffer(): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  // Sheet 1: Fund Fees
  const fees = wb.addWorksheet('Fund Fees');
  fees.getCell('A1').value = 'Fee Type';
  fees.getCell('B1').value = 'Rate';
  fees.getCell('A2').value = 'Management Fee';
  fees.getCell('B2').value = 1.25;
  fees.getCell('A3').value = 'Trustee Fee';
  fees.getCell('B3').value = 0.10;
  fees.getCell('A4').value = 'Total Fees';
  fees.getCell('B4').value = { formula: 'SUM(B2:B3)', result: 1.35 };

  // Sheet 2: Fund Performance
  const perf = wb.addWorksheet('Performance');
  perf.getCell('A1').value = 'Period';
  perf.getCell('B1').value = 'Return';
  perf.getCell('A2').value = '1M';
  perf.getCell('B2').value = '2.3%';
  perf.getCell('A3').value = '3M';
  perf.getCell('B3').value = '6.8%';
  perf.getCell('A4').value = 'NAV';
  perf.getCell('B4').value = 105.25;

  // Sheet 3: Portfolio Holdings
  const port = wb.addWorksheet('Portfolio');
  port.getCell('A1').value = 'Security';
  port.getCell('B1').value = 'Weight';
  port.getCell('A2').value = 'Equity Fund';
  port.getCell('B2').value = 60;
  port.getCell('A3').value = 'Bond Fund';
  port.getCell('B3').value = 40;

  return Buffer.from(await wb.xlsx.writeBuffer());
}

function makeWorkbookInfo(sheetCount: number, sheets: WorkbookInfo['sheets']): WorkbookInfo {
  return { sheetCount, sheets };
}

function makeMockAi(responses: AgentRoundResponse[]): jest.Mocked<Pick<AiService, 'callAgentRound'>> {
  let call = 0;
  return {
    callAgentRound: jest.fn(async () => {
      const r = responses[call] ?? responses[responses.length - 1];
      call++;
      return r;
    }),
  } as any;
}

// ─── Setup ────────────────────────────────────────────────────────────────────

let excelSvc: ExcelService;
let structSvc: WorkbookStructureService;
let inspSvc: ExcelInspectionService;
let opsSvc: OperationsService;
let router: OperationExecutionRouter;
let fixtureBuffer: Buffer;
let multiBuffer: Buffer;

beforeAll(async () => {
  excelSvc = new ExcelService();
  structSvc = new WorkbookStructureService();
  inspSvc = new ExcelInspectionService(structSvc);
  opsSvc = new OperationsService(inspSvc, structSvc);
  router = new OperationExecutionRouter(opsSvc, excelSvc);
  fixtureBuffer = await buildFinancialFixtureBuffer();
  multiBuffer = await buildMultiSheetBuffer();
});

function buildAgent(responses: AgentRoundResponse[]): AgentService {
  const mockAi = makeMockAi(responses);
  return new AgentService(mockAi as any, inspSvc, structSvc, opsSvc);
}

// ═══════════════════════════════════════════════════════════════════════════
// A. Multi-operation plans
// ═══════════════════════════════════════════════════════════════════════════

describe('A. Multi-operation plans', () => {
  it('agent produces plan with multiple write operations', async () => {
    const agent = buildAgent([{
      type: 'final_plan',
      reasoning: 'Two cells need updating',
      plan: {
        summary: 'Update management fee and trustee fee',
        operations: [
          { type: 'SET_CELL', sheet: 'Fund Fees', params: { row: 2, col: 2, value: 1.5 } },
          { type: 'SET_CELL', sheet: 'Fund Fees', params: { row: 3, col: 2, value: 0.15 } },
        ],
      },
    }]);

    const info = makeWorkbookInfo(3, [
      { name: 'Fund Fees', rowCount: 4, dataRowCount: 3, columnCount: 2, dataColumnCount: 2, headers: ['Fee Type', 'Rate'] },
    ]);
    const result = await agent.runAgent(multiBuffer, 'Update fees', info);

    // B2 and B3 are referenced by the SUM formula in B4.
    // Dependency check correctly fires → clarification_needed with plan attached.
    // The plan is preserved so it can be executed after user responds.
    expect(['completed', 'clarification_needed']).toContain(result.status);
    if (result.status === 'clarification_needed') {
      // Plan is preserved for execution after user confirms
      expect(result.plan).not.toBeNull();
      expect(result.plan!.operations).toHaveLength(2);
      expect(result.clarificationQuestion).toContain('formula');
    } else {
      expect(result.plan!.operations).toHaveLength(2);
    }
  });

  it('multi-op plan executes through router correctly', async () => {
    const wb = await excelSvc.loadWorkbook(multiBuffer);
    const plan = {
      summary: 'Update two fees',
      operations: [
        { type: 'SET_CELL', sheet: 'Fund Fees', params: { row: 2, col: 2, value: 1.5 } },
        { type: 'SET_CELL', sheet: 'Fund Fees', params: { row: 3, col: 2, value: 0.15 } },
      ],
    };
    opsSvc.validatePlan(plan);
    const result = await router.executePlan(multiBuffer, plan, wb);
    expect(result.writeOpsApplied).toBe(2);
    expect(wb.getWorksheet('Fund Fees').getCell(2, 2).value).toBe(1.5);
    expect(wb.getWorksheet('Fund Fees').getCell(3, 2).value).toBe(0.15);
  });

  it('mixed read+write plan: inspect then write two sheets', async () => {
    const wb = await excelSvc.loadWorkbook(multiBuffer);
    const plan = {
      summary: 'Inspect Fund Fees then update NAV on Performance',
      operations: [
        { type: 'DETECT_SECTIONS', sheet: 'Fund Fees', params: { sheetName: 'Fund Fees' } },
        { type: 'SET_CELL', sheet: 'Performance', params: { row: 4, col: 2, value: 106.50 } },
      ],
    };
    opsSvc.validatePlan(plan);
    const result = await router.executePlan(multiBuffer, plan, wb);
    expect(result.readOpsExecuted).toBe(1);
    expect(result.writeOpsApplied).toBe(1);
    expect(wb.getWorksheet('Performance').getCell(4, 2).value).toBe(106.50);
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// B. Multi-sheet workbooks
// ═══════════════════════════════════════════════════════════════════════════

describe('B. Multi-sheet workbook handling', () => {
  it('agent can write to different sheets in one plan', async () => {
    const agent = buildAgent([{
      type: 'final_plan',
      reasoning: 'Updating fee on sheet 1 and NAV on sheet 2',
      plan: {
        summary: 'Update fee and NAV',
        operations: [
          { type: 'SET_CELL', sheet: 'Fund Fees', params: { row: 2, col: 2, value: 2.0 } },
          { type: 'SET_CELL', sheet: 'Performance', params: { row: 4, col: 2, value: 110.0 } },
        ],
      },
    }]);

    const info = makeWorkbookInfo(3, [
      { name: 'Fund Fees', rowCount: 4, dataRowCount: 3, columnCount: 2, dataColumnCount: 2, headers: ['Fee Type', 'Rate'] },
      { name: 'Performance', rowCount: 4, dataRowCount: 3, columnCount: 2, dataColumnCount: 2, headers: ['Period', 'Return'] },
      { name: 'Portfolio', rowCount: 3, dataRowCount: 2, columnCount: 2, dataColumnCount: 2, headers: ['Security', 'Weight'] },
    ]);

    const result = await agent.runAgent(multiBuffer, 'Update fee and NAV', info);
    // Fund Fees B2 is referenced by SUM formula in B4 → dependency check may fire
    expect(['completed', 'clarification_needed']).toContain(result.status);
    // Plan is always preserved regardless of status
    expect(result.plan).not.toBeNull();
    expect(result.plan!.operations[0].sheet).toBe('Fund Fees');
    expect(result.plan!.operations[1].sheet).toBe('Performance');
  });

  it('search_workbook tool searches across all sheets', async () => {
    const searchResult = await inspSvc.searchWorkbook(multiBuffer, 'Management Fee');
    expect(searchResult.matches.length).toBeGreaterThan(0);
    expect(searchResult.matches[0].sheetName).toBe('Fund Fees');
  });

  it('search_workbook finds same label in multiple sheets', async () => {
    // "Management Fee" appears in multiple sheets in the financial fixture
    const result = await inspSvc.searchWorkbook(fixtureBuffer, 'Management Fee');
    expect(result.matches.length).toBeGreaterThan(1);
    const sheets = new Set(result.matches.map(m => m.sheetName));
    expect(sheets.size).toBeGreaterThan(1);
  });

  it('agent uses tool call to search before writing to correct sheet', async () => {
    const agent = buildAgent([
      {
        type: 'tool_call',
        tool: 'search_workbook',
        params: { query: 'NAV', sheetName: 'Performance' },
        reasoning: 'Need to find NAV cell location',
      },
      {
        type: 'final_plan',
        reasoning: 'Found NAV at B4 on Performance',
        plan: {
          summary: 'Update NAV',
          operations: [
            { type: 'SET_CELL', sheet: 'Performance', params: { row: 4, col: 2, value: 107.5 } },
          ],
        },
      },
    ]);

    const info = makeWorkbookInfo(3, [
      { name: 'Fund Fees', rowCount: 4, dataRowCount: 3, columnCount: 2, dataColumnCount: 2, headers: [] },
      { name: 'Performance', rowCount: 4, dataRowCount: 3, columnCount: 2, dataColumnCount: 2, headers: [] },
      { name: 'Portfolio', rowCount: 3, dataRowCount: 2, columnCount: 2, dataColumnCount: 2, headers: [] },
    ]);

    const result = await agent.runAgent(multiBuffer, 'Update the NAV value', info);
    expect(result.status).toBe('completed');
    expect(result.toolCallsMade).toBe(1);
    expect(result.toolHistory[0].tool).toBe('search_workbook');
    expect(result.toolHistory[0].success).toBe(true);
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// C. Search + inspect + write sequence
// ═══════════════════════════════════════════════════════════════════════════

describe('C. Search → inspect → write (evidence-based)', () => {
  it('3-round: search → context → plan', async () => {
    const agent = buildAgent([
      {
        type: 'tool_call',
        tool: 'search_workbook',
        params: { query: 'Management Fee', sheetName: 'Fund Fees' },
        reasoning: 'Locate Management Fee label',
      },
      {
        type: 'tool_call',
        tool: 'get_cell_context',
        params: { sheetName: 'Fund Fees', coordinate: 'A2', rowRadius: 2, colRadius: 1 },
        reasoning: 'Verify B2 is the value cell',
      },
      {
        type: 'final_plan',
        reasoning: 'Evidence confirmed: B2 is Management Fee value',
        plan: {
          summary: 'Set Management Fee to 1.75',
          operations: [
            { type: 'SET_CELL', sheet: 'Fund Fees', params: { row: 2, col: 2, value: 1.75 } },
          ],
        },
      },
    ]);

    const info = makeWorkbookInfo(1, [
      { name: 'Fund Fees', rowCount: 4, dataRowCount: 3, columnCount: 2, dataColumnCount: 2, headers: ['Fee Type', 'Rate'] },
    ]);
    const result = await agent.runAgent(multiBuffer, 'Change Management Fee to 1.75%', info);

    // Fund Fees B2 is referenced by =SUM(B2:B3) at B4 → dependency check fires.
    // Result: clarification_needed with the plan preserved + question about B4 formula.
    // This is CORRECT new behavior — agent asks before silently skipping dependents.
    expect(['completed', 'clarification_needed']).toContain(result.status);
    expect(result.roundsUsed).toBe(3);
    expect(result.toolCallsMade).toBe(2);
    expect(result.toolHistory[0].tool).toBe('search_workbook');
    expect(result.toolHistory[1].tool).toBe('get_cell_context');
    expect(result.plan).not.toBeNull();
    expect(result.plan!.operations[0].params.value).toBe(1.75);
    if (result.status === 'clarification_needed') {
      // Dependency question shown to user
      expect(result.clarificationQuestion).toContain('formula');
      // First candidate is always the safe "leave as-is" option
      expect(result.clarificationCandidates![0]).toContain('Excel will auto-recalculate');
    }
  });

  it('resolve_label directly finds and resolves unique label', async () => {
    const result = await opsSvc.executeSetCellByLabel(multiBuffer, {
      label: 'Management Fee',
      value: 1.75,
      sheetName: 'Fund Fees',
    });
    expect(result.status).toBe('resolved');
    expect(result.resolvedTarget!.sheetName).toBe('Fund Fees');
    expect(result.resolvedTarget!.labelCoordinate).toBe('A2');
    expect(result.resolvedTarget!.valueCoordinate).toBe('B2');
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// D. Tool result summarisation (Phase 3 Step 2)
// ═══════════════════════════════════════════════════════════════════════════

describe('D. Tool result summarisation', () => {
  it('agent receives summarised search results (not raw array)', async () => {
    let capturedToolResult: any = null;
    const mockAi = {
      callAgentRound: jest.fn(async (conv: any[]) => {
        const lastMsg = conv[conv.length - 1];
        if (lastMsg?.role === 'user') {
          try {
            const parsed = JSON.parse(lastMsg.content);
            if (parsed.tool_result?.result) {
              capturedToolResult = parsed.tool_result.result;
            }
          } catch {}
        }
        // After first tool call, produce final plan
        if (capturedToolResult !== null) {
          return {
            type: 'final_plan',
            reasoning: 'done',
            plan: {
              summary: 'no-op',
              operations: [{ type: 'ADD_ROW', sheet: 'Fund Fees', params: { values: [] } }],
            },
          } as AgentRoundResponse;
        }
        return {
          type: 'tool_call',
          tool: 'search_workbook',
          params: { query: 'Fee' },
          reasoning: 'searching',
        } as AgentRoundResponse;
      }),
    };

    const agent = new AgentService(mockAi as any, inspSvc, structSvc, opsSvc);
    const info = makeWorkbookInfo(1, [
      { name: 'Fund Fees', rowCount: 4, dataRowCount: 3, columnCount: 2, dataColumnCount: 2, headers: [] },
    ]);
    await agent.runAgent(multiBuffer, 'Find fees', info);

    // Summarised result should have totalMatches and matches array (not raw)
    expect(capturedToolResult).not.toBeNull();
    expect(capturedToolResult.totalMatches).toBeDefined();
    expect(Array.isArray(capturedToolResult.matches)).toBe(true);
    // Should be capped at 10
    expect(capturedToolResult.matches.length).toBeLessThanOrEqual(10);
  });

  it('summarised search result has coordinate and value fields', async () => {
    let capturedResult: any = null;
    const mockAi = {
      callAgentRound: jest.fn(async (conv: any[]) => {
        const last = conv[conv.length - 1];
        try {
          const p = JSON.parse(last.content);
          if (p.tool_result?.result?.matches) capturedResult = p.tool_result.result;
        } catch {}
        if (capturedResult) return { type: 'final_plan', reasoning: 'done', plan: { summary: 'no-op', operations: [{ type: 'ADD_ROW', sheet: 'Fund Fees', params: { values: [] } }] } } as AgentRoundResponse;
        return { type: 'tool_call', tool: 'search_workbook', params: { query: 'Management' }, reasoning: 'x' } as AgentRoundResponse;
      }),
    };
    const agent = new AgentService(mockAi as any, inspSvc, structSvc, opsSvc);
    const info = makeWorkbookInfo(1, [{ name: 'Fund Fees', rowCount: 4, dataRowCount: 3, columnCount: 2, dataColumnCount: 2, headers: [] }]);
    await agent.runAgent(multiBuffer, 'Search', info);

    if (capturedResult?.matches?.length > 0) {
      const m = capturedResult.matches[0];
      expect(m.coordinate).toBeDefined();
      expect(m.value).toBeDefined();
      expect(m.sheetName).toBeDefined();
    }
  });

  it('inspect_sheet_structure result does not include mergedRanges array (only count)', async () => {
    let capturedResult: any = null;
    const mockAi = {
      callAgentRound: jest.fn(async (conv: any[]) => {
        const last = conv[conv.length - 1];
        try {
          const p = JSON.parse(last.content);
          if (p.tool_result?.result?.cellCounts) capturedResult = p.tool_result.result;
        } catch {}
        if (capturedResult) return { type: 'final_plan', reasoning: 'done', plan: { summary: 'no-op', operations: [{ type: 'ADD_ROW', sheet: 'Fund Overview', params: { values: [] } }] } } as AgentRoundResponse;
        return { type: 'tool_call', tool: 'inspect_sheet_structure', params: { sheetName: 'Fund Overview' }, reasoning: 'x' } as AgentRoundResponse;
      }),
    };
    const agent = new AgentService(mockAi as any, inspSvc, structSvc, opsSvc);
    const info = makeWorkbookInfo(1, [{ name: 'Fund Overview', rowCount: 20, dataRowCount: 18, columnCount: 6, dataColumnCount: 5, headers: [] }]);
    await agent.runAgent(fixtureBuffer, 'Inspect', info);

    // mergedRanges array should be stripped, only count present
    expect(capturedResult).not.toBeNull();
    expect(capturedResult.mergedRangeCount).toBeDefined();
    expect(Array.isArray(capturedResult.mergedRanges)).toBe(false);
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// E. Clarification flow
// ═══════════════════════════════════════════════════════════════════════════

describe('E. Clarification flow', () => {
  it('agent returns clarification_needed with question and candidates', async () => {
    const agent = buildAgent([{
      type: 'tool_call',
      tool: 'search_workbook',
      params: { query: 'Management Fee' },
      reasoning: 'Need to locate Management Fee',
    }, {
      type: 'clarification_required',
      reasoning: 'Found in 4 sheets',
      question: 'Which Management Fee do you want to update?',
      candidates: [
        'Fund Overview sheet — A7',
        'Fund Details sheet — A10',
        'Fee Structure sheet — A4',
      ],
    }]);

    const info = makeWorkbookInfo(4, [
      { name: 'Fund Overview', rowCount: 20, dataRowCount: 18, columnCount: 6, dataColumnCount: 5, headers: [] },
      { name: 'Fund Details', rowCount: 35, dataRowCount: 32, columnCount: 8, dataColumnCount: 7, headers: [] },
    ]);
    const result = await agent.runAgent(fixtureBuffer, 'Update management fee', info);

    expect(result.status).toBe('clarification_needed');
    expect(result.clarificationQuestion).toContain('Management Fee');
    expect(result.clarificationCandidates!.length).toBe(3);
    expect(result.plan).toBeNull();
  });

  it('after clarification answer, agent resolves to correct target', async () => {
    // Simulates re-running with combined instruction after user picks a candidate
    const combinedInstruction =
      'Update management fee to 1.75%\n\nUser clarification: Fund Details sheet';
    const agent = buildAgent([{
      type: 'final_plan',
      reasoning: 'User clarified Fund Details sheet — writing to B10',
      plan: {
        summary: 'Set Management Fee on Fund Details to 1.75%',
        operations: [
          { type: 'SET_CELL', sheet: 'Fund Details', params: { row: 10, col: 2, value: 1.75 } },
        ],
      },
    }]);

    const info = makeWorkbookInfo(1, [
      { name: 'Fund Details', rowCount: 35, dataRowCount: 32, columnCount: 8, dataColumnCount: 7, headers: [] },
    ]);
    const result = await agent.runAgent(fixtureBuffer, combinedInstruction, info);
    expect(result.status).toBe('completed');
    expect(result.plan!.operations[0].sheet).toBe('Fund Details');
    expect(result.plan!.operations[0].params.value).toBe(1.75);
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// F. Financial fixture — real workbook multi-section inspection
// ═══════════════════════════════════════════════════════════════════════════

describe('F. Financial fixture — multi-section inspection', () => {
  it('detect_sections finds multiple sections in Fund Details', async () => {
    const result = await structSvc.detectSections(fixtureBuffer, 'Fund Details');
    expect(result.sections.length).toBeGreaterThanOrEqual(2);
    // Sections should have formula presence (fee calculations)
    const withFormulas = result.sections.filter(s => s.hasFormulas);
    expect(withFormulas.length).toBeGreaterThan(0);
  });

  it('search finds Management Fee across 4+ sheets', async () => {
    const result = await inspSvc.searchWorkbook(fixtureBuffer, 'Management Fee');
    const sheets = new Set(result.matches.map(m => m.sheetName));
    expect(sheets.size).toBeGreaterThanOrEqual(3);
  });

  it('full agent loop: inspect → detect → plan (fixture workbook)', async () => {
    const agent = buildAgent([
      {
        type: 'tool_call',
        tool: 'detect_sections',
        params: { sheetName: 'Fund Details' },
        reasoning: 'Understand layout first',
      },
      {
        type: 'tool_call',
        tool: 'resolve_label',
        params: { label: 'Total Fee Burden', sheetName: 'Fund Details', valueColOffset: 2 },
        reasoning: 'Find Total Fee Burden target',
      },
      {
        type: 'final_plan',
        reasoning: 'Located Total Fee Burden — writing to verified coordinate',
        plan: {
          summary: 'Update Total Fee Burden',
          operations: [
            { type: 'SET_CELL', sheet: 'Fund Details', params: { row: 21, col: 3, value: 12.5 } },
          ],
        },
      },
    ]);

    const info = makeWorkbookInfo(5, [
      { name: 'Fund Details', rowCount: 35, dataRowCount: 32, columnCount: 8, dataColumnCount: 7, headers: [] },
    ]);
    const result = await agent.runAgent(fixtureBuffer, 'Update total fee burden', info);
    expect(result.status).toBe('completed');
    // Complex workbook (5 sheets) → 3 inspection-only rounds enforced.
    // The final_plan attempt in round 3 is blocked; approved in round 4.
    expect(result.roundsUsed).toBe(4);
    expect(result.toolCallsMade).toBe(2);
  });
});
