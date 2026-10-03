/**
 * ai-prompt.spec.ts
 *
 * Phase 2 Step 4 — AI prompt content tests.
 *
 * These tests verify the content of the system prompt and user prompt
 * WITHOUT calling the real AI API. They ensure:
 *
 *   A. All 13 supported operations are documented in the system prompt
 *   B. Formula safety (allowFormulaOverwrite) is documented
 *   C. Inspect-first principle is present
 *   D. SET_CELL_BY_LABEL semantics are explained
 *   E. Read operation params are correct
 *   F. Hybrid operation semantics (ambiguous / not-found) are explained
 *   G. User prompt includes sheet count, headers, complexity hints
 *   H. Response parsing handles edge cases correctly (regression)
 *   I. Input validation still works (regression)
 *
 * Strategy: access private methods via (service as any).methodName()
 * to unit-test the prompt builders in isolation.
 */

import { ConfigService } from '@nestjs/config';
import { AiService } from '../ai.service';
import { OperationsService } from '../../operations/operations.service';
import { WorkbookStructureService } from '../../excel/workbook-structure.service';
import { ExcelInspectionService } from '../../excel/excel-inspection.service';
import { WorkbookInfo } from '../../excel/excel.service';

// ─── Setup ────────────────────────────────────────────────────────────────────

function buildAiService(): AiService {
  const structureSvc = new WorkbookStructureService();
  const inspectionSvc = new ExcelInspectionService(structureSvc);
  const opsSvc = new OperationsService(inspectionSvc, structureSvc);

  // Minimal ConfigService stub — no real API key needed for prompt tests
  const configStub = {
    get: (key: string) => {
      if (key === 'app.ai.apiKey') return 'test-key';
      if (key === 'app.ai.model') return 'claude-sonnet-4-6';
      if (key === 'app.ai.baseUrl') return 'https://api.anthropic.com';
      return undefined;
    },
  } as unknown as ConfigService;

  return new AiService(configStub, opsSvc);
}

function getSystemPrompt(svc: AiService): string {
  return (svc as any).buildSystemPrompt();
}

function getUserPrompt(svc: AiService, instruction: string, info: WorkbookInfo): string {
  return (svc as any).buildUserPrompt(instruction, info);
}

const simpleWorkbook: WorkbookInfo = {
  sheetCount: 1,
  sheets: [{
    name: 'Employees',
    rowCount: 5,
    dataRowCount: 4,
    columnCount: 3,
    dataColumnCount: 3,
    headers: ['Name', 'Department', 'Salary'],
  }],
};

const complexWorkbook: WorkbookInfo = {
  sheetCount: 4,
  sheets: [
    { name: 'Fund Overview', rowCount: 20, dataRowCount: 18, columnCount: 6, dataColumnCount: 5, headers: [] },
    { name: 'Fund Details', rowCount: 35, dataRowCount: 32, columnCount: 8, dataColumnCount: 7, headers: [] },
    { name: 'Portfolio', rowCount: 25, dataRowCount: 22, columnCount: 5, dataColumnCount: 5, headers: ['Security', 'Sector', 'Weight', 'Quantity', 'Value'] },
    { name: 'Fee Structure', rowCount: 12, dataRowCount: 10, columnCount: 4, dataColumnCount: 4, headers: [] },
  ],
};

let svc: AiService;
let prompt: string;

beforeAll(() => {
  svc = buildAiService();
  prompt = getSystemPrompt(svc);
});

// ═══════════════════════════════════════════════════════════════════════════
// A. All 13 operations present in system prompt
// ═══════════════════════════════════════════════════════════════════════════

describe('A. All 13 operations documented in system prompt', () => {

  describe('8 write operations', () => {
    const writeOps = [
      'SET_CELL',
      'SET_CELL_RANGE',
      'ADD_ROW',
      'DELETE_ROW',
      'ADD_FORMULA',
      'FORMAT_CELL',
      'ADD_COLUMN_HEADER',
      'SORT_SHEET',
    ];
    writeOps.forEach((op) => {
      it(`documents ${op}`, () => {
        expect(prompt).toContain(op);
      });
    });
  });

  describe('4 read (inspection) operations', () => {
    const readOps = [
      'INSPECT_WORKBOOK_STRUCTURE',
      'DETECT_HEADERS',
      'DETECT_SECTIONS',
      'GET_CELL_CONTEXT',
    ];
    readOps.forEach((op) => {
      it(`documents ${op}`, () => {
        expect(prompt).toContain(op);
      });
    });
  });

  it('documents SET_CELL_BY_LABEL (hybrid)', () => {
    expect(prompt).toContain('SET_CELL_BY_LABEL');
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// B. Formula safety documented
// ═══════════════════════════════════════════════════════════════════════════

describe('B. Formula safety in system prompt', () => {
  it('mentions allowFormulaOverwrite', () => {
    expect(prompt).toContain('allowFormulaOverwrite');
  });

  it('states formulas are protected by default', () => {
    expect(prompt.toLowerCase()).toMatch(/formula.*(protected|blocked|safe)/);
  });

  it('warns against setting allowFormulaOverwrite without explicit user request', () => {
    expect(prompt).toContain('NEVER set allowFormulaOverwrite');
  });

  it('mentions atomic / no partial writes for SET_CELL_RANGE', () => {
    expect(prompt.toLowerCase()).toMatch(/atomic|no partial/);
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// C. Inspect-first principle present
// ═══════════════════════════════════════════════════════════════════════════

describe('C. Inspect-first principle', () => {
  it('mentions inspecting before writing for complex workbooks', () => {
    expect(prompt.toLowerCase()).toMatch(/inspect.*(before|first)/);
  });

  it('states inspection operations are read-only / never modify', () => {
    expect(prompt.toLowerCase()).toMatch(/read.only|never modif/);
  });

  it('recommends DETECT_SECTIONS or DETECT_HEADERS for complex sheets', () => {
    expect(prompt).toContain('DETECT_SECTIONS');
    expect(prompt).toContain('DETECT_HEADERS');
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// D. SET_CELL_BY_LABEL semantics
// ═══════════════════════════════════════════════════════════════════════════

describe('D. SET_CELL_BY_LABEL semantics documented', () => {
  it('explains the label → value-cell lookup concept', () => {
    expect(prompt.toLowerCase()).toMatch(/label.*(search|find|lookup)/);
  });

  it('explains ambiguous match behavior (does not write)', () => {
    expect(prompt.toLowerCase()).toMatch(/ambiguous/);
    expect(prompt.toLowerCase()).toMatch(/does not write|does NOT write/i);
  });

  it('documents valueColOffset parameter', () => {
    expect(prompt).toContain('valueColOffset');
  });

  it('documents sheetName filter to narrow search', () => {
    expect(prompt).toContain('sheetName');
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// E. Read operation params documented
// ═══════════════════════════════════════════════════════════════════════════

describe('E. Read operation params', () => {
  it('GET_CELL_CONTEXT documents coordinate param', () => {
    expect(prompt).toContain('coordinate');
  });

  it('GET_CELL_CONTEXT documents rowRadius', () => {
    expect(prompt).toContain('rowRadius');
  });

  it('GET_CELL_CONTEXT documents colRadius', () => {
    expect(prompt).toContain('colRadius');
  });

  it('all inspection ops require sheetName', () => {
    expect(prompt).toContain('"sheetName"');
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// F. Examples present in system prompt
// ═══════════════════════════════════════════════════════════════════════════

describe('F. Examples in system prompt', () => {
  it('has a simple SET_CELL example', () => {
    expect(prompt).toContain('SET_CELL');
    expect(prompt).toContain('"row"');
    expect(prompt).toContain('"col"');
  });

  it('has a SET_CELL_BY_LABEL example', () => {
    expect(prompt).toContain('SET_CELL_BY_LABEL');
    expect(prompt).toContain('"label"');
  });

  it('has an inspect-first example (DETECT_SECTIONS before write)', () => {
    // Example 3 should show DETECT_SECTIONS + SET_CELL_BY_LABEL in sequence
    const detectIdx = prompt.indexOf('DETECT_SECTIONS');
    const labelIdx = prompt.lastIndexOf('SET_CELL_BY_LABEL');
    expect(detectIdx).toBeGreaterThan(-1);
    expect(labelIdx).toBeGreaterThan(detectIdx); // DETECT_SECTIONS before SET_CELL_BY_LABEL example
  });

  it('has an unsupported instruction example returning empty operations', () => {
    expect(prompt).toContain('"operations": []');
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// G. User prompt content
// ═══════════════════════════════════════════════════════════════════════════

describe('G. User prompt structure', () => {
  it('includes the instruction text', () => {
    const up = getUserPrompt(svc, 'Update management fee to 1.5%', simpleWorkbook);
    expect(up).toContain('Update management fee to 1.5%');
  });

  it('includes the sheet name', () => {
    const up = getUserPrompt(svc, 'Do something', simpleWorkbook);
    expect(up).toContain('Employees');
  });

  it('includes column headers', () => {
    const up = getUserPrompt(svc, 'Do something', simpleWorkbook);
    expect(up).toContain('Name');
    expect(up).toContain('Department');
    expect(up).toContain('Salary');
  });

  it('includes sheet count', () => {
    const up = getUserPrompt(svc, 'Do something', complexWorkbook);
    expect(up).toContain('4 sheet');
  });

  it('includes all sheet names for multi-sheet workbook', () => {
    const up = getUserPrompt(svc, 'Do something', complexWorkbook);
    expect(up).toContain('Fund Overview');
    expect(up).toContain('Fund Details');
    expect(up).toContain('Portfolio');
    expect(up).toContain('Fee Structure');
  });

  it('adds complexity hint for multi-sheet workbooks', () => {
    const up = getUserPrompt(svc, 'Do something', complexWorkbook);
    // Should mention inspection for complex workbooks
    expect(up.toLowerCase()).toMatch(/inspect|detect|complex/);
  });

  it('includes data row counts', () => {
    const up = getUserPrompt(svc, 'Do something', simpleWorkbook);
    expect(up).toContain('Data rows: 4');
  });

  it('shows "(none detected)" when sheet has no headers', () => {
    const up = getUserPrompt(svc, 'Do something', complexWorkbook);
    expect(up).toContain('(none detected)');
  });

  it('single-sheet workbook does not trigger complexity hint', () => {
    const up = getUserPrompt(svc, 'Do something', simpleWorkbook);
    // Should NOT say "consider inspection operations" for a simple 1-sheet workbook
    expect(up).not.toContain('NOTE: This workbook has multiple sheets');
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// H. JSON parsing regression
// ═══════════════════════════════════════════════════════════════════════════

describe('H. JSON parsing (regression)', () => {
  function parseAiResponse(raw: string) {
    return (svc as any).parseAiResponse(raw);
  }

  it('parses clean JSON', () => {
    const raw = JSON.stringify({ summary: 'test', operations: [{ type: 'SET_CELL', sheet: 'S', params: {} }] });
    const result = parseAiResponse(raw);
    expect(result.summary).toBe('test');
    expect(result.operations).toHaveLength(1);
  });

  it('strips json code fence', () => {
    const raw = '```json\n{"summary":"test","operations":[]}\n```';
    const result = parseAiResponse(raw);
    expect(result.summary).toBe('test');
  });

  it('strips plain code fence', () => {
    const raw = '```\n{"summary":"test","operations":[]}\n```';
    const result = parseAiResponse(raw);
    expect(result.summary).toBe('test');
  });

  it('throws on invalid JSON', () => {
    expect(() => parseAiResponse('not json at all')).toThrow();
  });

  it('throws when operations array is missing', () => {
    expect(() => parseAiResponse('{"summary":"test"}')).toThrow(/operations/i);
  });

  it('uses fallback summary when absent', () => {
    const raw = JSON.stringify({ operations: [] });
    const result = parseAiResponse(raw);
    expect(result.summary).toBe('No summary provided');
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// I. Input validation regression
// ═══════════════════════════════════════════════════════════════════════════

describe('I. Input validation (regression)', () => {
  function validateInputs(instruction: string, info: WorkbookInfo) {
    return (svc as any).validateInputs(instruction, info);
  }

  it('accepts valid instruction and workbook', () => {
    expect(() => validateInputs('Update fee', simpleWorkbook)).not.toThrow();
  });

  it('rejects empty instruction', () => {
    expect(() => validateInputs('', simpleWorkbook)).toThrow();
  });

  it('rejects whitespace-only instruction', () => {
    expect(() => validateInputs('   ', simpleWorkbook)).toThrow();
  });

  it('rejects instruction over 2000 chars', () => {
    expect(() => validateInputs('x'.repeat(2001), simpleWorkbook)).toThrow();
  });

  it('rejects workbook with no sheets', () => {
    expect(() => validateInputs('test', { sheetCount: 0, sheets: [] })).toThrow();
  });
});
