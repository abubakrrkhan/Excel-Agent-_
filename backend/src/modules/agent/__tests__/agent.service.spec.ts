/**
 * agent.service.spec.ts
 *
 * Phase 3 Step 1 — AgentService tests.
 *
 * Tests AgentService in isolation by mocking AiService.callAgentRound.
 * No real API calls. Uses real inspection services on real workbook buffers.
 *
 * Tests:
 *   A. Successful single-round plan production
 *   B. Multi-round: tool call → inspection result → final plan
 *   C. Clarification required when label is ambiguous
 *   D. Cannot proceed response from AI
 *   E. Max rounds safety limit
 *   F. Plan validation failure → AI self-correction round
 *   G. All 6 tools dispatch to correct services
 *   H. Conversation history grows correctly across rounds
 *   I. Activity log captures all agent steps
 *   J. Tool errors handled gracefully (success=false, not thrown)
 */

import * as ExcelJS from 'exceljs';
import { AgentService } from '../agent.service';
import { AiService } from '../../ai/ai.service';
import { ExcelInspectionService } from '../../excel/excel-inspection.service';
import { WorkbookStructureService } from '../../excel/workbook-structure.service';
import { OperationsService } from '../../operations/operations.service';
import { WorkbookInfo } from '../../excel/excel.service';
import {
  AgentRoundResponse,
  AgentExecutionResult,
} from '../agent.types';
import { buildFinancialFixtureBuffer } from '../../excel/fixtures/financial-workbook.fixture';

// ─── Helpers ──────────────────────────────────────────────────────────────────

async function buildSimpleBuffer(): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Fees');
  ws.getCell('A1').value = 'Fee Type';
  ws.getCell('B1').value = 'Rate';
  ws.getCell('A2').value = 'Management Fee Unique XYZ';
  ws.getCell('B2').value = 1.25;
  ws.getCell('A3').value = 'Admin Fee Unique ABC';
  ws.getCell('B3').value = 0.5;
  return Buffer.from(await wb.xlsx.writeBuffer());
}

const simpleWorkbookInfo: WorkbookInfo = {
  sheetCount: 1,
  sheets: [{
    name: 'Fees',
    rowCount: 3,
    dataRowCount: 3,
    columnCount: 2,
    dataColumnCount: 2,
    headers: ['Fee Type', 'Rate'],
  }],
};

// ─── Setup ────────────────────────────────────────────────────────────────────

let agentSvc: AgentService;
let mockAiCallAgentRound: jest.Mock;
let fixtureBuffer: Buffer;
let simpleBuffer: Buffer;

beforeAll(async () => {
  fixtureBuffer = await buildFinancialFixtureBuffer();
  simpleBuffer = await buildSimpleBuffer();
});

beforeEach(() => {
  const structureSvc = new WorkbookStructureService();
  const inspectionSvc = new ExcelInspectionService(structureSvc);
  const opsSvc = new OperationsService(inspectionSvc, structureSvc);

  // Mock AiService — only callAgentRound is used by AgentService
  mockAiCallAgentRound = jest.fn();
  const aiSvc = { callAgentRound: mockAiCallAgentRound } as unknown as AiService;

  agentSvc = new AgentService(aiSvc, inspectionSvc, structureSvc, opsSvc);
});

// ═══════════════════════════════════════════════════════════════════════════
// A. Single-round plan production
// ═══════════════════════════════════════════════════════════════════════════

describe('A. Single-round plan production', () => {
  it('returns completed when AI produces valid plan immediately', async () => {
    const validPlan = {
      summary: 'Set cell A1 to test',
      operations: [{ type: 'SET_CELL', sheet: 'Fees', params: { row: 1, col: 1, value: 'test' } }],
    };

    mockAiCallAgentRound.mockResolvedValueOnce({
      type: 'final_plan',
      reasoning: 'Target is clear',
      plan: validPlan,
    } as AgentRoundResponse);

    const result = await agentSvc.runAgent(simpleBuffer, 'Set A1 to test', simpleWorkbookInfo);

    expect(result.status).toBe('completed');
    expect(result.plan).toEqual(validPlan);
    expect(result.roundsUsed).toBe(1);
    expect(result.toolCallsMade).toBe(0);
  });

  it('plan is not null on completed status', async () => {
    mockAiCallAgentRound.mockResolvedValueOnce({
      type: 'final_plan',
      reasoning: 'done',
      plan: {
        summary: 'test',
        operations: [{ type: 'ADD_ROW', sheet: 'Fees', params: { values: ['X', 1] } }],
      },
    });

    const result = await agentSvc.runAgent(simpleBuffer, 'add a row', simpleWorkbookInfo);
    expect(result.status).toBe('completed');
    expect(result.plan).not.toBeNull();
    expect(result.plan!.operations).toHaveLength(1);
  });

  it('toolHistory is empty when no tools were called', async () => {
    mockAiCallAgentRound.mockResolvedValueOnce({
      type: 'final_plan',
      reasoning: 'obvious',
      plan: { summary: 'test', operations: [{ type: 'ADD_ROW', sheet: 'Fees', params: { values: [1] } }] },
    });

    const result = await agentSvc.runAgent(simpleBuffer, 'add row', simpleWorkbookInfo);
    expect(result.toolHistory).toHaveLength(0);
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// B. Multi-round: tool call → result → final plan
// ═══════════════════════════════════════════════════════════════════════════

describe('B. Multi-round: tool call then final plan', () => {
  it('calls tool, receives result, then produces final plan', async () => {
    // Round 1: AI calls search_workbook
    mockAiCallAgentRound.mockResolvedValueOnce({
      type: 'tool_call',
      tool: 'search_workbook',
      params: { query: 'Management Fee Unique XYZ', maxMatches: 5 },
      reasoning: 'Need to find the Management Fee cell',
    } as AgentRoundResponse);

    // Round 2: AI produces plan after seeing search results
    mockAiCallAgentRound.mockResolvedValueOnce({
      type: 'final_plan',
      reasoning: 'Found unique match at A2, value cell is B2',
      plan: {
        summary: 'Update Management Fee to 2.0',
        operations: [{ type: 'SET_CELL', sheet: 'Fees', params: { row: 2, col: 2, value: 2.0 } }],
      },
    } as AgentRoundResponse);

    const result = await agentSvc.runAgent(simpleBuffer, 'Update management fee to 2.0', simpleWorkbookInfo);

    expect(result.status).toBe('completed');
    expect(result.roundsUsed).toBe(2);
    expect(result.toolCallsMade).toBe(1);
    expect(result.toolHistory).toHaveLength(1);
    expect(result.toolHistory[0].tool).toBe('search_workbook');
    expect(result.toolHistory[0].success).toBe(true);
    expect(mockAiCallAgentRound).toHaveBeenCalledTimes(2);
  });

  it('conversation grows with each round', async () => {
    const capturedConversations: Array<Array<{ role: string; content: string }>> = [];

    mockAiCallAgentRound.mockImplementation(async (conv) => {
      capturedConversations.push([...conv]);
      if (capturedConversations.length === 1) {
        return {
          type: 'tool_call',
          tool: 'detect_sections',
          params: { sheetName: 'Fees' },
          reasoning: 'Inspect structure',
        };
      }
      return {
        type: 'final_plan',
        reasoning: 'done',
        plan: { summary: 'test', operations: [{ type: 'ADD_ROW', sheet: 'Fees', params: { values: [1] } }] },
      };
    });

    await agentSvc.runAgent(simpleBuffer, 'do something', simpleWorkbookInfo);

    // Round 1: system + user = 2 messages
    expect(capturedConversations[0].length).toBe(2);
    // Round 2: system + user + assistant (tool call) + user (tool result) = 4 messages
    expect(capturedConversations[1].length).toBe(4);
  });

  it('tool result is fed back to AI in next round', async () => {
    const capturedRound2: Array<{ role: string; content: string }> = [];

    mockAiCallAgentRound.mockImplementation(async (conv) => {
      if (conv.length === 2) {
        return {
          type: 'tool_call',
          tool: 'detect_headers',
          params: { sheetName: 'Fees' },
          reasoning: 'need headers',
        };
      }
      capturedRound2.push(...conv);
      return {
        type: 'final_plan',
        reasoning: 'done',
        plan: { summary: 't', operations: [{ type: 'ADD_ROW', sheet: 'Fees', params: { values: [1] } }] },
      };
    });

    await agentSvc.runAgent(simpleBuffer, 'inspect then write', simpleWorkbookInfo);

    // The last user message in round 2 should contain tool_result
    const toolResultMsg = capturedRound2.find(
      (m) => m.role === 'user' && m.content.includes('tool_result'),
    );
    expect(toolResultMsg).toBeDefined();
    const parsed = JSON.parse(toolResultMsg!.content);
    expect(parsed.tool_result.tool).toBe('detect_headers');
    expect(parsed.tool_result.success).toBe(true);
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// C. Clarification required
// ═══════════════════════════════════════════════════════════════════════════

describe('C. Clarification required', () => {
  it('returns clarification_needed with question and candidates', async () => {
    mockAiCallAgentRound.mockResolvedValueOnce({
      type: 'clarification_required',
      reasoning: 'Management Fee found in 4 locations',
      question: 'Which Management Fee do you want to update?',
      candidates: [
        'Fund Overview sheet, cell A7',
        'Fund Details sheet, cell A10',
        'Fee Structure sheet, cell A4',
      ],
    } as AgentRoundResponse);

    const result = await agentSvc.runAgent(fixtureBuffer, 'Update management fee', {
      sheetCount: 4,
      sheets: [{ name: 'Fund Overview', rowCount: 20, dataRowCount: 18, columnCount: 6, dataColumnCount: 5, headers: [] }],
    });

    expect(result.status).toBe('clarification_needed');
    expect(result.plan).toBeNull();
    expect(result.clarificationQuestion).toContain('Management Fee');
    expect(result.clarificationCandidates).toHaveLength(3);
    expect(result.roundsUsed).toBe(1);
  });

  it('plan is null on clarification_needed', async () => {
    mockAiCallAgentRound.mockResolvedValueOnce({
      type: 'clarification_required',
      reasoning: 'ambiguous',
      question: 'Which one?',
      candidates: ['A', 'B'],
    });

    const result = await agentSvc.runAgent(simpleBuffer, 'update fee', simpleWorkbookInfo);
    expect(result.plan).toBeNull();
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// D. Cannot proceed
// ═══════════════════════════════════════════════════════════════════════════

describe('D. Cannot proceed', () => {
  it('returns cannot_proceed with reason', async () => {
    mockAiCallAgentRound.mockResolvedValueOnce({
      type: 'cannot_proceed',
      reasoning: 'instruction requires deleting the entire workbook',
      reason: 'This operation is not supported — cannot delete the workbook.',
    } as AgentRoundResponse);

    const result = await agentSvc.runAgent(simpleBuffer, 'delete everything', simpleWorkbookInfo);

    expect(result.status).toBe('cannot_proceed');
    expect(result.plan).toBeNull();
    expect(result.cannotProceedReason).toContain('not supported');
  });

  it('plan is null and toolHistory preserved on cannot_proceed', async () => {
    // Round 1: calls a tool
    mockAiCallAgentRound.mockResolvedValueOnce({
      type: 'tool_call',
      tool: 'detect_sections',
      params: { sheetName: 'Fees' },
      reasoning: 'inspecting',
    });
    // Round 2: cannot proceed after inspection
    mockAiCallAgentRound.mockResolvedValueOnce({
      type: 'cannot_proceed',
      reasoning: 'not possible',
      reason: 'Cannot fulfil this instruction.',
    });

    const result = await agentSvc.runAgent(simpleBuffer, 'impossible task', simpleWorkbookInfo);

    expect(result.status).toBe('cannot_proceed');
    expect(result.plan).toBeNull();
    expect(result.toolHistory).toHaveLength(1); // tool was called before cannot_proceed
    expect(result.toolCallsMade).toBe(1);
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// E. Max rounds safety limit
// ═══════════════════════════════════════════════════════════════════════════

describe('E. Max rounds safety limit', () => {
  it('returns max_rounds_reached when AI keeps calling tools', async () => {
    // AI always calls a tool, never produces a plan
    mockAiCallAgentRound.mockResolvedValue({
      type: 'tool_call',
      tool: 'detect_sections',
      params: { sheetName: 'Fees' },
      reasoning: 'still looking',
    } as AgentRoundResponse);

    const result = await agentSvc.runAgent(
      simpleBuffer,
      'inspect forever',
      simpleWorkbookInfo,
      { maxRounds: 3, maxToolCalls: 10 },
    );

    expect(result.status).toBe('max_rounds_reached');
    expect(result.roundsUsed).toBe(3);
    expect(result.plan).toBeNull();
  });

  it('returns max_rounds_reached when tool call limit is hit', async () => {
    mockAiCallAgentRound.mockResolvedValue({
      type: 'tool_call',
      tool: 'detect_headers',
      params: { sheetName: 'Fees' },
      reasoning: 'still inspecting',
    } as AgentRoundResponse);

    const result = await agentSvc.runAgent(
      simpleBuffer,
      'inspect too much',
      simpleWorkbookInfo,
      { maxRounds: 20, maxToolCalls: 2 },
    );

    expect(result.status).toBe('max_rounds_reached');
    expect(result.toolCallsMade).toBe(2);
    expect(result.plan).toBeNull();
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// F. Plan validation failure → self-correction round
// ═══════════════════════════════════════════════════════════════════════════

describe('F. Validation failure self-correction', () => {
  it('gives AI another round when plan fails validation', async () => {
    // Round 1: AI produces invalid plan (bad op type)
    mockAiCallAgentRound.mockResolvedValueOnce({
      type: 'final_plan',
      reasoning: 'first attempt',
      plan: {
        summary: 'bad plan',
        operations: [{ type: 'INVALID_OP_TYPE', sheet: 'Fees', params: {} }],
      },
    } as AgentRoundResponse);

    // Round 2: AI produces valid plan after seeing validation error
    mockAiCallAgentRound.mockResolvedValueOnce({
      type: 'final_plan',
      reasoning: 'corrected',
      plan: {
        summary: 'valid plan',
        operations: [{ type: 'ADD_ROW', sheet: 'Fees', params: { values: ['X', 1] } }],
      },
    } as AgentRoundResponse);

    const result = await agentSvc.runAgent(simpleBuffer, 'add a row', simpleWorkbookInfo);

    expect(result.status).toBe('completed');
    expect(result.roundsUsed).toBe(2);
    expect(result.plan!.summary).toBe('valid plan');
    expect(mockAiCallAgentRound).toHaveBeenCalledTimes(2);
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// G. Tool dispatch — all 6 tools execute correctly
// ═══════════════════════════════════════════════════════════════════════════

describe('G. Tool dispatch to correct services', () => {
  async function runSingleToolTest(
    tool: string,
    params: Record<string, any>,
    buffer: Buffer,
    workbookInfo: WorkbookInfo,
  ): Promise<AgentExecutionResult> {
    mockAiCallAgentRound.mockResolvedValueOnce({
      type: 'tool_call',
      tool,
      params,
      reasoning: 'testing tool dispatch',
    });
    mockAiCallAgentRound.mockResolvedValueOnce({
      type: 'cannot_proceed',
      reasoning: 'done testing',
      reason: 'Test complete',
    });

    return agentSvc.runAgent(buffer, 'test tool', workbookInfo, { maxRounds: 5, maxToolCalls: 5 });
  }

  it('search_workbook dispatches and returns matches', async () => {
    const result = await runSingleToolTest(
      'search_workbook',
      { query: 'Management Fee', maxMatches: 5 },
      fixtureBuffer,
      { sheetCount: 1, sheets: [{ name: 'Fund Overview', rowCount: 20, dataRowCount: 18, columnCount: 6, dataColumnCount: 5, headers: [] }] },
    );
    expect(result.toolHistory[0].tool).toBe('search_workbook');
    expect(result.toolHistory[0].success).toBe(true);
    expect(result.toolHistory[0].result).toHaveProperty('matches');
  });

  it('detect_sections dispatches and returns sections', async () => {
    const result = await runSingleToolTest(
      'detect_sections',
      { sheetName: 'Fund Details' },
      fixtureBuffer,
      { sheetCount: 1, sheets: [{ name: 'Fund Details', rowCount: 30, dataRowCount: 28, columnCount: 5, dataColumnCount: 4, headers: [] }] },
    );
    expect(result.toolHistory[0].tool).toBe('detect_sections');
    expect(result.toolHistory[0].success).toBe(true);
    expect(result.toolHistory[0].result).toHaveProperty('sections');
  });

  it('detect_headers dispatches and returns candidates', async () => {
    const result = await runSingleToolTest(
      'detect_headers',
      { sheetName: 'Fund Details' },
      fixtureBuffer,
      { sheetCount: 1, sheets: [{ name: 'Fund Details', rowCount: 30, dataRowCount: 28, columnCount: 5, dataColumnCount: 4, headers: [] }] },
    );
    expect(result.toolHistory[0].success).toBe(true);
    expect(result.toolHistory[0].result).toHaveProperty('candidates');
  });

  it('get_cell_context dispatches and returns context', async () => {
    const result = await runSingleToolTest(
      'get_cell_context',
      { sheetName: 'Fund Overview', coordinate: 'A5', rowRadius: 2, colRadius: 2 },
      fixtureBuffer,
      { sheetCount: 1, sheets: [{ name: 'Fund Overview', rowCount: 20, dataRowCount: 18, columnCount: 6, dataColumnCount: 5, headers: [] }] },
    );
    expect(result.toolHistory[0].success).toBe(true);
    expect(result.toolHistory[0].result).toHaveProperty('targetCell');
  });

  it('inspect_sheet_structure dispatches and returns structure', async () => {
    const result = await runSingleToolTest(
      'inspect_sheet_structure',
      { sheetName: 'Fund Overview' },
      fixtureBuffer,
      { sheetCount: 1, sheets: [{ name: 'Fund Overview', rowCount: 20, dataRowCount: 18, columnCount: 6, dataColumnCount: 5, headers: [] }] },
    );
    expect(result.toolHistory[0].success).toBe(true);
    expect(result.toolHistory[0].result).toHaveProperty('cellCounts');
  });

  it('resolve_label dispatches and returns resolution', async () => {
    const result = await runSingleToolTest(
      'resolve_label',
      { label: 'Management Fee Unique XYZ', sheetName: 'Fees' },
      simpleBuffer,
      simpleWorkbookInfo,
    );
    expect(result.toolHistory[0].tool).toBe('resolve_label');
    expect(result.toolHistory[0].success).toBe(true);
    expect(result.toolHistory[0].result).toHaveProperty('status');
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// H. Activity log
// ═══════════════════════════════════════════════════════════════════════════

describe('H. Activity log', () => {
  it('activity log contains starting message', async () => {
    mockAiCallAgentRound.mockResolvedValueOnce({
      type: 'final_plan',
      reasoning: 'clear',
      plan: { summary: 'test', operations: [{ type: 'ADD_ROW', sheet: 'Fees', params: { values: [1] } }] },
    });

    const result = await agentSvc.runAgent(simpleBuffer, 'add row', simpleWorkbookInfo);
    expect(result.activityLog.length).toBeGreaterThan(0);
    expect(result.activityLog[0]).toContain('Starting agent');
  });

  it('activity log records each round', async () => {
    mockAiCallAgentRound.mockResolvedValueOnce({
      type: 'tool_call',
      tool: 'detect_sections',
      params: { sheetName: 'Fees' },
      reasoning: 'inspect',
    });
    mockAiCallAgentRound.mockResolvedValueOnce({
      type: 'final_plan',
      reasoning: 'done',
      plan: { summary: 'test', operations: [{ type: 'ADD_ROW', sheet: 'Fees', params: { values: [1] } }] },
    });

    const result = await agentSvc.runAgent(simpleBuffer, 'test', simpleWorkbookInfo);
    const roundMessages = result.activityLog.filter((m) => m.includes('Round'));
    expect(roundMessages.length).toBe(2);
  });

  it('activity log records tool calls', async () => {
    mockAiCallAgentRound.mockResolvedValueOnce({
      type: 'tool_call',
      tool: 'search_workbook',
      params: { query: 'Fee', maxMatches: 3 },
      reasoning: 'search',
    });
    mockAiCallAgentRound.mockResolvedValueOnce({
      type: 'cannot_proceed',
      reasoning: 'done',
      reason: 'test',
    });

    const result = await agentSvc.runAgent(simpleBuffer, 'find fee', simpleWorkbookInfo);
    const toolMessages = result.activityLog.filter((m) => m.includes('Tool call'));
    expect(toolMessages.length).toBe(1);
    expect(toolMessages[0]).toContain('search_workbook');
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// I. Tool errors handled gracefully
// ═══════════════════════════════════════════════════════════════════════════

describe('I. Tool errors handled gracefully', () => {
  it('non-existent sheet name causes tool success=false, not thrown', async () => {
    mockAiCallAgentRound.mockResolvedValueOnce({
      type: 'tool_call',
      tool: 'detect_sections',
      params: { sheetName: 'NoSuchSheetXYZ' },
      reasoning: 'testing error handling',
    });
    mockAiCallAgentRound.mockResolvedValueOnce({
      type: 'cannot_proceed',
      reasoning: 'sheet not found',
      reason: 'Sheet does not exist.',
    });

    const result = await agentSvc.runAgent(simpleBuffer, 'inspect missing sheet', simpleWorkbookInfo);

    expect(result.toolHistory[0].success).toBe(false);
    expect(result.toolHistory[0].error).toBeDefined();
    expect(result.toolHistory[0].error).toMatch(/not found/i);
    // Agent continues despite tool error
    expect(result.status).toBe('cannot_proceed');
  });

  it('unknown tool name causes success=false, not thrown', async () => {
    mockAiCallAgentRound.mockResolvedValueOnce({
      type: 'tool_call',
      tool: 'nonexistent_tool' as any,
      params: {},
      reasoning: 'testing unknown tool',
    });
    mockAiCallAgentRound.mockResolvedValueOnce({
      type: 'cannot_proceed',
      reasoning: 'unknown tool',
      reason: 'Test',
    });

    const result = await agentSvc.runAgent(simpleBuffer, 'test', simpleWorkbookInfo);
    expect(result.toolHistory[0].success).toBe(false);
    expect(result.toolHistory[0].error).toContain('Unknown tool');
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// J. Default config
// ═══════════════════════════════════════════════════════════════════════════

describe('J. Default agent config', () => {
  it('uses DEFAULT_AGENT_CONFIG when no config is passed', async () => {
    // AI always tools — we need it to hit MAX rounds
    mockAiCallAgentRound.mockResolvedValue({
      type: 'tool_call',
      tool: 'detect_headers',
      params: { sheetName: 'Fees' },
      reasoning: 'looping',
    } as AgentRoundResponse);

    const result = await agentSvc.runAgent(simpleBuffer, 'loop', simpleWorkbookInfo);
    // Default maxRounds=8, maxToolCalls=12 → rounds limit hits first
    expect(result.status).toBe('max_rounds_reached');
    expect(result.roundsUsed).toBe(8);
  });
});
