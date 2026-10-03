import {
  Injectable,
  Logger,
  BadRequestException,
  InternalServerErrorException,
  ServiceUnavailableException,
} from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import axios, { AxiosError } from 'axios';
import { WorkbookInfo } from '../excel/excel.service';
import { OperationsService } from '../operations/operations.service';

// ─── Shared types (used by OperationsService + AiController) ─────────────────

export interface ExcelOperation {
  type: string;
  sheet: string;
  params: Record<string, any>;
}

export interface OperationPlan {
  operations: ExcelOperation[];
  summary: string;
}

// ─── Internal raw AI response shape ──────────────────────────────────────────

interface RawAiPlan {
  summary?: unknown;
  operations?: unknown;
}

// ─── Service ──────────────────────────────────────────────────────────────────

const AI_TIMEOUT_MS = 30_000; // 30 seconds
const MAX_INSTRUCTION_LENGTH = 2000;

@Injectable()
export class AiService {
  private readonly logger = new Logger(AiService.name);

  constructor(
    private readonly configService: ConfigService,
    private readonly operationsService: OperationsService,
  ) {}

  /**
   * Generate a validated operation plan from a natural-language instruction.
   *
   * Flow:
   *   1. Validate inputs
   *   2. Call AI provider
   *   3. Extract + parse JSON from AI response
   *   4. Validate plan through OperationsService
   *   5. Return accepted plan
   *
   * The AI never touches the Excel file.
   * Excel modification happens in Day 4.
   */
  async generateOperationPlan(
    instruction: string,
    workbookInfo: WorkbookInfo,
  ): Promise<OperationPlan> {
    // ── 1. Validate inputs ────────────────────────────────────────────────────
    this.validateInputs(instruction, workbookInfo);

    // ── 2. Read provider config ───────────────────────────────────────────────
    const apiKey = this.configService.get<string>('app.ai.apiKey');
    const model = this.configService.get<string>('app.ai.model');
    const baseUrl = this.configService.get<string>('app.ai.baseUrl');

    if (!apiKey || apiKey === 'your_api_key_here') {
      throw new ServiceUnavailableException(
        'AI provider is not configured. Set AI_API_KEY in environment variables.',
      );
    }

    this.logger.log(
      `Generating operation plan | model=${model} | instruction="${instruction.slice(0, 80)}..."`,
    );

    // ── 3. Call AI provider ───────────────────────────────────────────────────
    const rawText = await this.callAiProvider({
      apiKey,
      model,
      baseUrl,
      systemPrompt: this.buildSystemPrompt(),
      userPrompt: this.buildUserPrompt(instruction, workbookInfo),
    });

    // ── 4. Parse JSON ─────────────────────────────────────────────────────────
    const plan = this.parseAiResponse(rawText);

    // ── 5. Validate through OperationsService ─────────────────────────────────
    // This is the security gate — unsupported types, missing params etc. rejected here
    this.operationsService.validatePlan(plan);

    this.logger.log(
      `Plan accepted: ${plan.operations.length} operation(s) — ${plan.summary}`,
    );

    return plan;
  }

  /**
   * Make a single agent reasoning round with full conversation history.
   *
   * Used by AgentService for multi-round tool-calling.
   * Sends the full conversation (system + all prior messages) to the AI
   * and returns a parsed AgentRoundResponse.
   *
   * The response is parsed as JSON and returned as-is — AgentService
   * decides what to do based on the `type` field.
   */
  async callAgentRound(
    conversation: Array<{ role: string; content: string }>,
  ): Promise<import('../agent/agent.types').AgentRoundResponse> {
    const apiKey = this.configService.get<string>('app.ai.apiKey');
    const model = this.configService.get<string>('app.ai.model');
    const baseUrl = this.configService.get<string>('app.ai.baseUrl');

    if (!apiKey || apiKey === 'your_api_key_here') {
      throw new ServiceUnavailableException(
        'AI provider is not configured. Set AI_API_KEY in environment variables.',
      );
    }

    // Extract system message and user/assistant turns
    const systemMsg = conversation.find((m) => m.role === 'system');
    const turns = conversation
      .filter((m) => m.role !== 'system')
      .map((m) => ({ role: m.role as 'user' | 'assistant', content: m.content }));

    if (turns.length === 0) {
      throw new BadRequestException('Agent conversation must have at least one user message');
    }

    try {
      const response = await axios.post(
        `${baseUrl}/v1/messages`,
        {
          model,
          max_tokens: 4096,
          system: systemMsg?.content ?? '',
          messages: turns,
        },
        {
          headers: {
            'x-api-key': apiKey,
            'anthropic-version': '2023-06-01',
            'content-type': 'application/json',
          },
          timeout: AI_TIMEOUT_MS,
        },
      );

      const raw: string = response.data?.content?.[0]?.text ?? '';
      if (!raw) {
        throw new InternalServerErrorException('AI agent returned empty response');
      }

      return this.parseAgentRound(raw);
    } catch (err) {
      // Re-throw NestJS HTTP exceptions as-is
      if (
        err instanceof BadRequestException ||
        err instanceof ServiceUnavailableException ||
        err instanceof InternalServerErrorException
      ) {
        throw err;
      }
      const axiosErr = err as import('axios').AxiosError;
      if (axiosErr.code === 'ECONNABORTED' || axiosErr.code === 'ETIMEDOUT') {
        throw new ServiceUnavailableException('AI agent request timed out');
      }
      if (axiosErr.response?.status === 401) {
        throw new ServiceUnavailableException('AI provider authentication failed');
      }
      throw new ServiceUnavailableException(
        `AI agent call failed: ${axiosErr.message ?? 'unknown error'}`,
      );
    }
  }

  /**
   * Parse a raw AI agent response string into an AgentRoundResponse.
   * The AI is instructed to respond with JSON only.
   */
  private parseAgentRound(
    raw: string,
  ): import('../agent/agent.types').AgentRoundResponse {
    const cleaned = raw
      .replace(/^```(?:json)?[\r\n]*/i, '')
      .replace(/[\r\n]*```\s*$/i, '')
      .trim();

    let parsed: Record<string, any>;
    try {
      parsed = JSON.parse(cleaned);
    } catch {
      this.logger.error('Agent returned non-JSON response', raw.slice(0, 200));
      // Treat unparseable responses as cannot_proceed
      return {
        type: 'cannot_proceed',
        reasoning: 'AI returned non-JSON response',
        reason: 'AI returned an unrecognised response format. Please try again.',
      };
    }

    const type = parsed.type;
    if (!type || !['tool_call', 'final_plan', 'clarification_required', 'cannot_proceed'].includes(type)) {
      return {
        type: 'cannot_proceed',
        reasoning: `Unknown response type: ${type}`,
        reason: 'AI returned an unknown response type. Please try again.',
      };
    }

    return parsed as import('../agent/agent.types').AgentRoundResponse;
  }

  private validateInputs(instruction: string, workbookInfo: WorkbookInfo): void {
    if (!instruction || typeof instruction !== 'string') {
      throw new BadRequestException('Instruction must be a non-empty string');
    }

    const trimmed = instruction.trim();
    if (trimmed.length === 0) {
      throw new BadRequestException('Instruction cannot be empty or whitespace only');
    }

    if (trimmed.length > MAX_INSTRUCTION_LENGTH) {
      throw new BadRequestException(
        `Instruction too long. Maximum is ${MAX_INSTRUCTION_LENGTH} characters.`,
      );
    }

    if (!workbookInfo || typeof workbookInfo !== 'object') {
      throw new BadRequestException('Workbook context is required');
    }

    if (!Array.isArray(workbookInfo.sheets) || workbookInfo.sheets.length === 0) {
      throw new BadRequestException(
        'Workbook context must contain at least one sheet',
      );
    }
  }

  // ─── Private: AI provider call ────────────────────────────────────────────

  private async callAiProvider(params: {
    apiKey: string;
    model: string;
    baseUrl: string;
    systemPrompt: string;
    userPrompt: string;
  }): Promise<string> {
    try {
      const response = await axios.post(
        `${params.baseUrl}/v1/messages`,
        {
          model: params.model,
          max_tokens: 2048,
          system: params.systemPrompt,
          messages: [{ role: 'user', content: params.userPrompt }],
        },
        {
          headers: {
            'x-api-key': params.apiKey,
            'anthropic-version': '2023-06-01',
            'content-type': 'application/json',
          },
          timeout: AI_TIMEOUT_MS,
        },
      );

      const text: string = response.data?.content?.[0]?.text ?? '';
      if (!text) {
        throw new InternalServerErrorException(
          'AI provider returned an empty response',
        );
      }

      return text;
    } catch (err) {
      // Re-throw NestJS HTTP exceptions as-is
      if (
        err instanceof BadRequestException ||
        err instanceof ServiceUnavailableException ||
        err instanceof InternalServerErrorException
      ) {
        throw err;
      }

      const axiosErr = err as AxiosError;

      if (axiosErr.code === 'ECONNABORTED' || axiosErr.code === 'ETIMEDOUT') {
        this.logger.error('AI provider request timed out');
        throw new ServiceUnavailableException(
          'AI provider request timed out. Please try again.',
        );
      }

      if (axiosErr.response) {
        const status = axiosErr.response.status;
        this.logger.error(
          `AI provider HTTP error: ${status}`,
          JSON.stringify(axiosErr.response.data),
        );

        if (status === 401) {
          throw new ServiceUnavailableException(
            'AI provider authentication failed. Check AI_API_KEY.',
          );
        }
        if (status === 429) {
          throw new ServiceUnavailableException(
            'AI provider rate limit exceeded. Please try again later.',
          );
        }
        throw new ServiceUnavailableException(
          `AI provider returned an error (HTTP ${status})`,
        );
      }

      // Network / DNS / connection error
      this.logger.error('AI provider connection error', axiosErr.message);
      throw new ServiceUnavailableException(
        'Could not reach AI provider. Check network and AI_BASE_URL.',
      );
    }
  }

  // ─── Private: JSON parsing ────────────────────────────────────────────────

  private parseAiResponse(raw: string): OperationPlan {
    // Strip markdown code fences — AI sometimes wraps JSON in ```json ... ```
    const cleaned = raw
      .replace(/^```(?:json)?[\r\n]*/i, '')
      .replace(/[\r\n]*```\s*$/i, '')
      .trim();

    let parsed: RawAiPlan;
    try {
      parsed = JSON.parse(cleaned);
    } catch {
      this.logger.error(
        'AI returned non-JSON response',
        `First 200 chars: ${raw.slice(0, 200)}`,
      );
      throw new BadRequestException(
        'AI returned an invalid response — could not parse JSON. ' +
          'This is an AI reliability issue; please try again.',
      );
    }

    // Structural checks before passing to OperationsService
    if (!parsed || typeof parsed !== 'object') {
      throw new BadRequestException('AI response is not a JSON object');
    }

    if (!Array.isArray(parsed.operations)) {
      throw new BadRequestException(
        'AI response missing required "operations" array',
      );
    }

    // summary is optional but normalise it
    const summary =
      typeof parsed.summary === 'string' && parsed.summary.trim()
        ? parsed.summary.trim()
        : 'No summary provided';

    return {
      summary,
      operations: parsed.operations as ExcelOperation[],
    };
  }

  // ─── Private: prompts ─────────────────────────────────────────────────────

  private buildSystemPrompt(): string {
    return `You are an intelligent Excel workbook operation planner for complex financial and fund-report workbooks.

Your job is to convert a natural-language instruction into a precise, safe JSON operation plan.

ABSOLUTE OUTPUT RULES:
1. Return ONLY a single valid JSON object. Nothing else.
2. Do NOT use Markdown. Do NOT wrap JSON in code fences.
3. Do NOT add explanatory text outside the JSON.
4. Do NOT invent operation types. Use ONLY the types listed below.
5. You are generating a PLAN — you are NOT applying changes yourself.
6. If the instruction cannot be fulfilled with supported operations, return an empty operations array with an explanatory summary.

REQUIRED JSON STRUCTURE:
{
  "summary": "One sentence describing what this plan will do",
  "operations": [
    {
      "type": "OPERATION_TYPE",
      "sheet": "Exact sheet name as it appears in the workbook context",
      "params": { ...operation-specific fields }
    }
  ]
}

INSPECT-FIRST PRINCIPLE (IMPORTANT FOR COMPLEX WORKBOOKS):
For financial/fund reports with multiple sheets, merged headers, or repeated labels,
place inspection operations BEFORE write operations in your plan.
Inspection operations are read-only and never modify the workbook.
They let you verify the exact target before writing.

Recommended sequence for complex instructions:
  1. DETECT_SECTIONS or DETECT_HEADERS — understand worksheet layout
  2. GET_CELL_CONTEXT — verify the exact target cell and its surroundings
  3. Then write using verified coordinates

For simple, clearly identified targets, inspection is optional.

FORMULA SAFETY (CRITICAL):
Financial workbooks contain formulas (SUM, AVERAGE, cross-sheet references).
The system BLOCKS any write that would overwrite a formula cell unless explicitly authorized.
- Default: allowFormulaOverwrite is absent or false — formula cells are PROTECTED.
- To replace a formula (only when user explicitly requests it): allowFormulaOverwrite: true
- NEVER set allowFormulaOverwrite: true unless the user explicitly asks to replace a formula.
- When uncertain whether a target cell contains a formula, use GET_CELL_CONTEXT first.

CELL ADDRESSING:
- All row and col values are 1-based integers (row 1 = first row, col 1 = column A).
- Use sheet names EXACTLY as they appear in the workbook context.
- Data rows typically begin at row 2 (row 1 = headers).

SUPPORTED OPERATION CATALOG:

── READ OPERATIONS (safe, never modify the workbook) ──

INSPECT_WORKBOOK_STRUCTURE — get sheet dimensions, merged cells, hidden rows/cols, cell type counts
  params: { "sheetName": "<exact sheet name>" }

DETECT_HEADERS — identify candidate header rows in a sheet
  params: { "sheetName": "<exact sheet name>" }

DETECT_SECTIONS — identify table/section regions separated by blank rows
  params: { "sheetName": "<exact sheet name>" }

GET_CELL_CONTEXT — inspect a cell and its surrounding rows/columns
  params: {
    "sheetName": "<exact sheet name>",
    "coordinate": "<A1-style address, e.g. G42>",
    "rowRadius": <optional integer 1-20, default 3>,
    "colRadius": <optional integer 1-20, default 3>
  }

── HYBRID OPERATION (reads first, writes only on unambiguous match) ──

SET_CELL_BY_LABEL — find a value cell by its row label, then write to the cell beside it
  Use when you know the label text but not the exact coordinate.
  If the label is found in multiple locations, the operation returns 'ambiguous' and does NOT write.
  Use sheetName to narrow the search to one sheet when the label may repeat.
  params: {
    "label": "<label text to search for>",
    "value": <new value>,
    "sheetName": "<optional: restrict search to this sheet>",
    "valueColOffset": <optional integer, default 1 = one column right of label>,
    "allowFormulaOverwrite": <optional boolean, default false>
  }

── WRITE OPERATIONS ──

SET_CELL — set a single cell value
  params: {
    "row": <integer>, "col": <integer>, "value": <string|number|boolean|null>,
    "allowFormulaOverwrite": <optional boolean, default false>
  }

SET_CELL_RANGE — set the same value across a rectangular range
  Entire range is pre-scanned for formulas before any cell is written (atomic — no partial writes).
  params: {
    "startRow": <integer>, "startCol": <integer>, "endRow": <integer>, "endCol": <integer>,
    "value": <any>,
    "allowFormulaOverwrite": <optional boolean, default false>
  }

ADD_ROW — append a new row of values to the end of a sheet
  params: { "values": [<value1>, <value2>, ...] }

DELETE_ROW — remove a specific row (1-based); rows above shift up
  params: { "row": <integer> }

ADD_FORMULA — write an Excel formula to a cell
  params: { "row": <integer>, "col": <integer>, "formula": "<e.g. SUM(B2:B10)>" }

FORMAT_CELL — apply visual formatting (does not change cell value)
  params: {
    "row": <integer>, "col": <integer>,
    "style": { "font": {...}, "fill": {...}, "alignment": {...}, "numFmt": "..." }
  }

ADD_COLUMN_HEADER — set a label in row 1 of a column
  params: { "col": <integer>, "header": "<header text>" }

SORT_SHEET — sort all data rows (rows 2+) by a column; row 1 preserved as header
  params: { "col": <integer>, "order": "asc" | "desc" }

EXAMPLES:

Example 1 — Simple write (target is clear from context):
Instruction: "Set Alice salary to 100000" (Name=col1, Salary=col3, Alice in row 2)
{ "summary": "Set salary to 100000 for Alice in row 2", "operations": [{ "type": "SET_CELL", "sheet": "Employees", "params": { "row": 2, "col": 3, "value": 100000 } }] }

Example 2 — Label-based update for financial report:
Instruction: "Change Management Fee to 1.5%"
{ "summary": "Update Management Fee to 1.5% using label search", "operations": [{ "type": "SET_CELL_BY_LABEL", "sheet": "Fund Details", "params": { "label": "Management Fee", "value": "1.5%", "sheetName": "Fund Details", "valueColOffset": 1 } }] }

Example 3 — Inspect first, then write for complex workbook:
Instruction: "Update the NAV per unit in the quarterly report"
{ "summary": "Inspect Fund Overview then update NAV per unit", "operations": [{ "type": "DETECT_SECTIONS", "sheet": "Fund Overview", "params": { "sheetName": "Fund Overview" } }, { "type": "SET_CELL_BY_LABEL", "sheet": "Fund Overview", "params": { "label": "NAV per Unit", "value": 105.50, "sheetName": "Fund Overview", "valueColOffset": 1 } }] }

Example 4 — Unsupported instruction:
Instruction: "Email this file to finance@company.com"
{ "summary": "Cannot send emails — file operations only. No changes made.", "operations": [] }`;
  }

  private buildUserPrompt(
    instruction: string,
    workbookInfo: WorkbookInfo,
  ): string {
    const sheetCount = workbookInfo.sheets.length;
    const sheetSummaries = workbookInfo.sheets
      .map((sheet, idx) => {
        const headerLine =
          sheet.headers.length > 0
            ? `  Headers (row 1): ${sheet.headers.map((h, i) => `col${i + 1}="${h}"`).join(', ')}`
            : '  Headers (row 1): (none detected)';
        const complexity =
          sheet.dataRowCount > 50 || sheet.dataColumnCount > 10
            ? ' [COMPLEX — consider using inspection operations first]'
            : '';
        return [
          `Sheet ${idx + 1} of ${sheetCount}: "${sheet.name}"${complexity}`,
          `  Data rows: ${sheet.dataRowCount}`,
          `  Columns: ${sheet.dataColumnCount}`,
          headerLine,
        ].join('\n');
      })
      .join('\n\n');

    const complexityHint =
      sheetCount > 2
        ? '\nNOTE: This workbook has multiple sheets. Consider using DETECT_SECTIONS or GET_CELL_CONTEXT before writing to complex sheets.'
        : '';

    return `WORKBOOK CONTEXT (${sheetCount} sheet${sheetCount !== 1 ? 's' : ''}):
${sheetSummaries}${complexityHint}

USER INSTRUCTION:
${instruction.trim()}

Generate the JSON operation plan now. Return ONLY the JSON object.`;
  }
}
