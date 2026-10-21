/**
 * agent.service.ts
 *
 * Phase 3 Step 1 — AI Agent Core.
 *
 * Orchestrates a multi-round tool-calling loop:
 *
 *   Round 1: AI receives workbook summary + available tools
 *            → AI calls a tool (e.g. detect_sections)
 *   Round 2: Tool result fed back to AI
 *            → AI calls another tool OR produces final plan
 *   ...
 *   Final:   AI emits final_plan, clarification_required, or cannot_proceed
 *
 * Architecture boundaries:
 *   AgentService      → orchestration only
 *   AiService         → all LLM calls (via callAiRound)
 *   ExcelInspectionService / WorkbookStructureService → tool execution
 *   OperationsService → plan validation (security gate, unchanged)
 *   OperationExecutionRouter → plan execution (unchanged)
 */

import { Injectable, Logger } from '@nestjs/common';
import { AiService } from '../ai/ai.service';
import { ExcelInspectionService } from '../excel/excel-inspection.service';
import { WorkbookStructureService } from '../excel/workbook-structure.service';
import { OperationsService } from '../operations/operations.service';
import { WorkbookInfo } from '../excel/excel.service';
import {
  AgentConfig,
  AgentExecutionResult,
  AgentMessage,
  AgentRoundResponse,
  AgentToolName,
  AgentToolResult,
  DEFAULT_AGENT_CONFIG,
  COMPLEX_AGENT_CONFIG,
  INSPECTION_ONLY_ROUNDS,
  isComplexWorkbook,
} from './agent.types';

@Injectable()
export class AgentService {
  private readonly logger = new Logger(AgentService.name);

  constructor(
    private readonly aiService: AiService,
    private readonly inspectionService: ExcelInspectionService,
    private readonly structureService: WorkbookStructureService,
    private readonly operationsService: OperationsService,
  ) {}

  // ─── Public API ────────────────────────────────────────────────────────────

  /**
   * Run the agent loop for a given instruction and workbook.
   *
   * @param buffer        The workbook buffer (read-only — never modified here)
   * @param instruction   The user's natural-language instruction
   * @param workbookInfo  Pre-computed workbook summary (from ExcelService.inspectWorkbook)
   * @param config        Optional: override max rounds / tool calls
   */
  async runAgent(
    buffer: Buffer,
    instruction: string,
    workbookInfo: WorkbookInfo,
    config?: Partial<AgentConfig>,
  ): Promise<AgentExecutionResult> {
    // Auto-select config: complex workbooks get more rounds and a mandatory
    // inspection-only phase to reduce hallucination on unfamiliar structures.
    const autoConfig = isComplexWorkbook(workbookInfo) ? COMPLEX_AGENT_CONFIG : DEFAULT_AGENT_CONFIG;
    const cfg: AgentConfig = {
      maxRounds: config?.maxRounds ?? autoConfig.maxRounds,
      maxToolCalls: config?.maxToolCalls ?? autoConfig.maxToolCalls,
    };
    const isComplex = isComplexWorkbook(workbookInfo);

    const activityLog: string[] = [];
    const toolHistory: AgentToolResult[] = [];
    const conversation: AgentMessage[] = [];

    // ── Evidence Registry (anti-hallucination) ────────────────────────────
    // Track every coordinate + sheet discovered by tool results.
    // Before accepting a final_plan, verify all write targets exist here.
    const discoveredCoordinates = new Set<string>(); // format: "SheetName::A1"
    const discoveredSheets = new Set<string>();

    const registerEvidence = (toolName: string, result: Record<string, any>) => {
      // Extract coordinates from tool results and add to evidence registry
      if (toolName === 'search_workbook') {
        for (const m of result.matches ?? []) {
          if (m.sheetName && m.coordinate) {
            discoveredCoordinates.add(`${m.sheetName}::${m.coordinate}`);
            discoveredSheets.add(m.sheetName);
          }
        }
      } else if (toolName === 'get_cell_context') {
        if (result.sheetName) discoveredSheets.add(result.sheetName);
        if (result.targetCell?.coordinate && result.sheetName) {
          discoveredCoordinates.add(`${result.sheetName}::${result.targetCell.coordinate}`);
        }
        // Register all cells in context
        for (const c of result.cells ?? []) {
          if (c.coordinate && result.sheetName) {
            discoveredCoordinates.add(`${result.sheetName}::${c.coordinate}`);
          }
        }
        for (const c of result.rowLabels ?? []) {
          if (c.coordinate && result.sheetName) {
            discoveredCoordinates.add(`${result.sheetName}::${c.coordinate}`);
          }
        }
        for (const c of result.colLabels ?? []) {
          if (c.coordinate && result.sheetName) {
            discoveredCoordinates.add(`${result.sheetName}::${c.coordinate}`);
          }
        }
      } else if (toolName === 'resolve_label') {
        if (result.resolvedTarget?.sheetName && result.resolvedTarget?.valueCoordinate) {
          discoveredCoordinates.add(`${result.resolvedTarget.sheetName}::${result.resolvedTarget.valueCoordinate}`);
          discoveredSheets.add(result.resolvedTarget.sheetName);
        }
        if (result.resolvedTarget?.sheetName && result.resolvedTarget?.labelCoordinate) {
          discoveredCoordinates.add(`${result.resolvedTarget.sheetName}::${result.resolvedTarget.labelCoordinate}`);
        }
        for (const c of result.candidates ?? []) {
          if (c.sheetName && c.valueCoordinate) {
            discoveredCoordinates.add(`${c.sheetName}::${c.valueCoordinate}`);
            discoveredSheets.add(c.sheetName);
          }
          if (c.sheetName && c.labelCoordinate) {
            discoveredCoordinates.add(`${c.sheetName}::${c.labelCoordinate}`);
          }
        }
      } else if (toolName === 'detect_sections') {
        if (result.sheetName) discoveredSheets.add(result.sheetName);
        for (const s of result.sections ?? []) {
          if (s.startCoordinate && result.sheetName) {
            discoveredCoordinates.add(`${result.sheetName}::${s.startCoordinate}`);
          }
        }
      } else if (toolName === 'detect_headers' || toolName === 'inspect_sheet_structure') {
        if (result.sheetName) discoveredSheets.add(result.sheetName);
      }
    };

    const log = (msg: string) => {
      this.logger.log(`[Agent] ${msg}`);
      activityLog.push(msg);
    };

    log(`Starting agent for instruction: "${instruction.slice(0, 80)}"`);
    log(`Workbook: ${workbookInfo.sheetCount} sheet(s) — ${isComplex ? 'COMPLEX (14 rounds, 3 inspection-only)' : 'standard (8 rounds)'}`);

    // Initialise conversation with system prompt + initial user message
    conversation.push({
      role: 'system',
      content: this.buildAgentSystemPrompt(),
    });
    conversation.push({
      role: 'user',
      content: this.buildInitialUserMessage(instruction, workbookInfo),
    });

    let roundsUsed = 0;
    let toolCallsMade = 0;

    // ── Main agent loop ──────────────────────────────────────────────────────
    while (roundsUsed < cfg.maxRounds) {
      roundsUsed++;
      log(`Round ${roundsUsed}/${cfg.maxRounds} — calling AI...`);

      // Ask AI what to do next
      let roundResponse: AgentRoundResponse;
      try {
        roundResponse = await this.aiService.callAgentRound(conversation);
      } catch (err) {
        log(`AI call failed in round ${roundsUsed}: ${err.message}`);
        return this.buildResult('cannot_proceed', {
          roundsUsed, toolCallsMade, toolHistory, activityLog,
          cannotProceedReason: `AI provider error: ${err.message}`,
        });
      }

      // ── Dispatch on AI decision ────────────────────────────────────────────

      if (roundResponse.type === 'cannot_proceed') {
        log(`Agent cannot proceed: ${roundResponse.reason}`);
        return this.buildResult('cannot_proceed', {
          roundsUsed, toolCallsMade, toolHistory, activityLog,
          cannotProceedReason: roundResponse.reason,
        });
      }

      if (roundResponse.type === 'clarification_required') {
        log(`Agent needs clarification: ${roundResponse.question}`);
        return this.buildResult('clarification_needed', {
          roundsUsed, toolCallsMade, toolHistory, activityLog,
          clarificationQuestion: roundResponse.question,
          clarificationCandidates: roundResponse.candidates,
        });
      }

      if (roundResponse.type === 'final_plan') {
        // Enforce inspection-only phase: complex workbooks must spend the first
        // INSPECTION_ONLY_ROUNDS rounds gathering evidence before any writes.
        if (isComplex && roundsUsed <= INSPECTION_ONLY_ROUNDS) {
          log(`Round ${roundsUsed}: complex workbook — inspection-only phase, rejecting early plan`);
          conversation.push({
            role: 'assistant',
            content: JSON.stringify({ type: 'final_plan', plan: roundResponse.plan }),
          });
          conversation.push({
            role: 'user',
            content: `This is a complex multi-sheet workbook. You are still in the inspection phase (round ${roundsUsed} of ${INSPECTION_ONLY_ROUNDS} required inspection rounds). Do NOT produce a plan yet. Call more inspection tools to build higher confidence before writing. Continue inspecting.`,
          });
          continue;
        }

        // ── ANTI-HALLUCINATION GATE 1: Require at least one tool call ─────
        if (toolCallsMade === 0) {
          log(`Plan rejected: no tool calls made before final plan — forcing inspection`);
          conversation.push({
            role: 'assistant',
            content: JSON.stringify({ type: 'final_plan', plan: roundResponse.plan }),
          });
          conversation.push({
            role: 'user',
            content: `REJECTED: You produced a plan without calling ANY inspection tools first. You MUST call search_workbook or get_cell_context to verify cell locations before producing a plan. Start by searching for the key terms in the instruction.`,
          });
          continue;
        }

        // ── ANTI-HALLUCINATION GATE 2: Coordinate evidence check ──────────
        // Verify every write target's sheet was actually inspected
        const writeOps = roundResponse.plan.operations.filter(
          op => ['SET_CELL', 'SET_CELL_RANGE', 'ADD_ROW', 'DELETE_ROW', 'ADD_FORMULA', 'FORMAT_CELL', 'ADD_COLUMN_HEADER'].includes(op.type)
        );
        const unverifiedOps: string[] = [];
        for (const op of writeOps) {
          if (op.sheet && !discoveredSheets.has(op.sheet)) {
            unverifiedOps.push(`${op.type} on sheet "${op.sheet}" — this sheet was never inspected`);
          }
          // For SET_CELL, verify the specific coordinate was seen
          if (op.type === 'SET_CELL' && op.params?.row && op.params?.col && op.sheet) {
            let colLetter = '';
            let c = op.params.col;
            while (c > 0) {
              colLetter = String.fromCharCode(64 + (c % 26 || 26)) + colLetter;
              c = Math.floor((c - 1) / 26);
            }
            const coord = `${op.sheet}::${colLetter}${op.params.row}`;
            if (!discoveredCoordinates.has(coord)) {
              unverifiedOps.push(`SET_CELL at ${colLetter}${op.params.row} on "${op.sheet}" — this coordinate was never returned by any tool`);
            }
          }
        }

        if (unverifiedOps.length > 0) {
          log(`Plan rejected: ${unverifiedOps.length} unverified coordinate(s) — forcing re-inspection`);
          conversation.push({
            role: 'assistant',
            content: JSON.stringify({ type: 'final_plan', plan: roundResponse.plan }),
          });
          conversation.push({
            role: 'user',
            content: `REJECTED — HALLUCINATION DETECTED: The following operations reference coordinates that were NOT returned by any tool in this session:\n${unverifiedOps.map(u => `• ${u}`).join('\n')}\n\nYou MUST call search_workbook or get_cell_context to discover the actual coordinates before producing a plan. Do NOT guess or assume coordinates. Call a tool now.`,
          });
          continue;
        }

        log(`Agent produced final plan: ${roundResponse.plan.summary}`);
        log(`Evidence check passed — ${discoveredCoordinates.size} verified coordinate(s)`);
        log(`Validating plan (${roundResponse.plan.operations.length} operation(s))...`);

        try {
          this.operationsService.validatePlan(roundResponse.plan);
        } catch (err) {
          log(`Plan validation failed: ${err.message} — asking AI to revise`);
          // Feed the validation error back so AI can self-correct
          conversation.push({
            role: 'assistant',
            content: JSON.stringify({ type: 'final_plan', plan: roundResponse.plan }),
          });
          conversation.push({
            role: 'user',
            content: `The plan failed validation: ${err.message}. Please revise the plan to fix this issue.`,
          });
          continue; // another round
        }

        // ── ANTI-HALLUCINATION GATE 3: Post-plan verification ─────────────
        // For each SET_CELL target, call get_cell_context to verify the cell
        // actually contains what the AI expects (current value, not formula, etc.)
        const setCellOps = roundResponse.plan.operations.filter(op => op.type === 'SET_CELL');
        if (setCellOps.length > 0 && setCellOps.length <= 5) {
          let verificationFailed = false;
          const verificationErrors: string[] = [];

          for (const op of setCellOps) {
            if (!op.params?.row || !op.params?.col || !op.sheet) continue;
            let colLetter = '';
            let c = op.params.col;
            while (c > 0) {
              colLetter = String.fromCharCode(64 + (c % 26 || 26)) + colLetter;
              c = Math.floor((c - 1) / 26);
            }
            const coordinate = `${colLetter}${op.params.row}`;
            try {
              const cellCtx = await this.structureService.getCellContext(
                buffer, op.sheet, coordinate, { rowRadius: 1, colRadius: 1 },
              );
              // Check if target cell has a formula that shouldn't be overwritten
              if (cellCtx.targetCell?.type === 'formula' && !op.params.allowFormulaOverwrite) {
                verificationFailed = true;
                verificationErrors.push(
                  `${coordinate} on "${op.sheet}" contains formula "${cellCtx.targetCell.formula}" but allowFormulaOverwrite is not set`,
                );
              }
            } catch {
              // Non-fatal — if verification fails, log but allow (the sheet/cell might not exist)
              verificationErrors.push(`Could not verify ${coordinate} on "${op.sheet}" — sheet or coordinate may be invalid`);
              verificationFailed = true;
            }
          }

          if (verificationFailed) {
            log(`Post-plan verification found issues: ${verificationErrors.join('; ')}`);
            conversation.push({
              role: 'assistant',
              content: JSON.stringify({ type: 'final_plan', plan: roundResponse.plan }),
            });
            conversation.push({
              role: 'user',
              content: `POST-PLAN VERIFICATION FAILED:\n${verificationErrors.map(e => `• ${e}`).join('\n')}\n\nRevise your plan to address these issues. If a cell contains a formula, ask the user before overwriting. If a coordinate is invalid, re-inspect to find the correct one.`,
            });
            continue;
          }
          log(`Post-plan verification passed for ${setCellOps.length} SET_CELL operation(s)`);
        }

        log(`Plan validated successfully — ${roundResponse.plan.operations.length} operation(s)`);

        // ── Dependency check ──────────────────────────────────────────────────
        // Scan the workbook for cells whose formulas reference the cells being
        // written. If found, ask the user — they may want those updated too.
        // Note: Excel recalculates formulas automatically on open, so this is
        // purely informational. User decides whether to update or leave them.
        const writeTargets = roundResponse.plan.operations
          .filter(op => ['SET_CELL', 'SET_CELL_RANGE'].includes(op.type))
          .map(op => {
            const col = op.params.col;
            const row = op.params.row;
            if (!col || !row) return null;
            // Convert col number to letter (1=A, 2=B, ...)
            let colLetter = '';
            let c = col;
            while (c > 0) {
              colLetter = String.fromCharCode(64 + (c % 26 || 26)) + colLetter;
              c = Math.floor((c - 1) / 26);
            }
            return { sheetName: op.sheet, coordinate: `${colLetter}${row}` };
          })
          .filter((t): t is { sheetName: string; coordinate: string } => t !== null);

        if (writeTargets.length > 0) {
          try {
            const dependents = await this.inspectionService.findDependentCells(
              buffer, writeTargets,
            );
            if (dependents.length > 0) {
              log(`Found ${dependents.length} dependent formula cell(s) — asking user`);
              const depCandidates = dependents.map(
                d => `${d.sheetName} → ${d.coordinate}  [formula: =${d.formula}]  depends on: ${d.dependsOn}`,
              );
              return this.buildResult('clarification_needed', {
                roundsUsed, toolCallsMade, toolHistory, activityLog,
                plan: roundResponse.plan,
                clarificationQuestion:
                  `Your change affects cells that contain formulas referencing ` +
                  `the updated cell(s). Excel will recalculate these automatically ` +
                  `when the file is opened — but if you want to override any formula ` +
                  `with a fixed value, select it below:`,
                clarificationCandidates: [
                  `Leave all formulas as-is — Excel will auto-recalculate on open ✓`,
                  ...depCandidates,
                ],
              });
            }
          } catch (depErr) {
            // Non-fatal — dependency scan failed, proceed without it
            log(`Dependency scan skipped: ${depErr.message}`);
          }
        }

        return this.buildResult('completed', {
          roundsUsed, toolCallsMade, toolHistory, activityLog,
          plan: roundResponse.plan,
        });
      }

      // ── Tool call ──────────────────────────────────────────────────────────
      if (roundResponse.type === 'tool_call') {
        if (toolCallsMade >= cfg.maxToolCalls) {
          log(`Tool call limit (${cfg.maxToolCalls}) reached before final plan`);
          return this.buildResult('max_rounds_reached', {
            roundsUsed, toolCallsMade, toolHistory, activityLog,
          });
        }

        log(`Tool call: ${roundResponse.tool}(${JSON.stringify(roundResponse.params).slice(0, 100)})`);
        if (roundResponse.reasoning) {
          log(`  Reason: ${roundResponse.reasoning.slice(0, 120)}`);
        }

        // Execute the tool
        const toolResult = await this.executeTool(buffer, roundResponse.tool, roundResponse.params);
        toolHistory.push(toolResult);
        toolCallsMade++;

        // Register discovered evidence from successful tool results
        if (toolResult.success) {
          registerEvidence(roundResponse.tool, toolResult.result);
        }

        if (toolResult.success) {
          log(`  → Tool succeeded`);
        } else {
          log(`  → Tool failed: ${toolResult.error}`);
        }

        // Append AI's tool call to conversation
        conversation.push({
          role: 'assistant',
          content: JSON.stringify({
            type: 'tool_call',
            tool: roundResponse.tool,
            params: roundResponse.params,
            reasoning: roundResponse.reasoning,
          }),
        });

        // Summarise large tool results before adding to context.
        // This prevents token-limit issues on large financial workbooks.
        const summarised = toolResult.success
          ? this.summariseToolResult(roundResponse.tool, toolResult.result)
          : null;

        // Append tool result to conversation so AI can reason over it
        conversation.push({
          role: 'user',
          content: JSON.stringify({
            tool_result: {
              tool: toolResult.tool,
              success: toolResult.success,
              result: summarised,
              error: toolResult.error ?? null,
            },
            remaining_rounds: cfg.maxRounds - roundsUsed,
            remaining_tool_calls: cfg.maxToolCalls - toolCallsMade,
          }),
        });

        continue; // next round
      }
    }

    // Exceeded max rounds without conclusion
    log(`Max rounds (${cfg.maxRounds}) reached without final plan`);
    return this.buildResult('max_rounds_reached', {
      roundsUsed, toolCallsMade, toolHistory, activityLog,
    });
  }

  // ─── Private: tool execution ──────────────────────────────────────────────

  private async executeTool(
    buffer: Buffer,
    tool: AgentToolName,
    params: Record<string, any>,
  ): Promise<AgentToolResult> {
    try {
      let result: Record<string, any>;

      switch (tool) {
        case 'inspect_sheet_structure':
          result = await this.structureService.inspectSheetStructure(
            buffer,
            params.sheetName,
          );
          break;

        case 'detect_headers':
          result = await this.structureService.detectHeaders(
            buffer,
            params.sheetName,
          );
          break;

        case 'detect_sections':
          result = await this.structureService.detectSections(
            buffer,
            params.sheetName,
          );
          break;

        case 'get_cell_context':
          result = await this.structureService.getCellContext(
            buffer,
            params.sheetName,
            params.coordinate,
            {
              rowRadius: params.rowRadius,
              colRadius: params.colRadius,
            },
          );
          break;

        case 'search_workbook':
          result = await this.inspectionService.searchWorkbook(
            buffer,
            params.query,
            {
              sheetName: params.sheetName,
              caseInsensitive: params.caseInsensitive ?? true,
              maxMatches: params.maxMatches ?? 10,
              includeContext: params.includeContext ?? true,
            },
          );
          break;

        case 'resolve_label':
          result = await this.operationsService.executeSetCellByLabel(
            buffer,
            params,
          );
          break;

        default:
          return {
            tool,
            params,
            result: {},
            success: false,
            error: `Unknown tool: ${tool}`,
          };
      }

      return { tool, params, result: result as Record<string, any>, success: true };
    } catch (err) {
      return {
        tool,
        params,
        result: {},
        success: false,
        error: err.message ?? 'Tool execution failed',
      };
    }
  }

  // ─── Private: tool result summarisation ──────────────────────────────────

  /**
   * Summarise large tool results before adding to the AI conversation.
   *
   * Prevents token-limit failures on large financial workbooks:
   *   - search_workbook: cap at MAX_SEARCH_MATCHES, strip raw cell arrays
   *   - detect_sections: cap at MAX_SECTIONS, return only key fields
   *   - detect_headers: return only high/medium confidence candidates
   *   - inspect_sheet_structure: strip cell-level arrays, keep counts + bounds
   *   - get_cell_context: keep target cell + up to MAX_CONTEXT_CELLS neighbours
   *   - resolve_label: pass through as-is (small result)
   *
   * Full raw results are preserved in toolHistory for inspection.
   */
  private summariseToolResult(
    tool: AgentToolName,
    result: Record<string, any>,
  ): Record<string, any> {
    const MAX_SEARCH_MATCHES = 10;
    const MAX_SECTIONS = 8;
    const MAX_CONTEXT_CELLS = 16;

    switch (tool) {
      case 'search_workbook': {
        const matches: any[] = result.matches ?? [];
        const truncated = matches.length > MAX_SEARCH_MATCHES;
        return {
          query: result.query,
          totalMatches: matches.length,
          truncated,
          matches: matches.slice(0, MAX_SEARCH_MATCHES).map((m: any) => ({
            sheetName: m.sheetName,
            coordinate: m.coordinate,
            value: m.value,
            type: m.type,
            // Include only essential context (not raw cell arrays)
            context: m.context ? {
              rowLabels: (m.context.rowLabels ?? []).slice(0, 3).map((c: any) => ({
                coordinate: c.coordinate,
                value: c.value,
              })),
              nearbyAbove: (m.context.nearbyAbove ?? []).slice(0, 2).map((c: any) => ({
                coordinate: c.coordinate,
                value: c.value,
              })),
            } : undefined,
          })),
          note: truncated
            ? `Showing ${MAX_SEARCH_MATCHES} of ${matches.length} matches. Use sheetName param to narrow search.`
            : undefined,
        };
      }

      case 'detect_sections': {
        const sections: any[] = result.sections ?? [];
        const truncated = sections.length > MAX_SECTIONS;
        return {
          sheetName: result.sheetName,
          totalSections: sections.length,
          sections: sections.slice(0, MAX_SECTIONS).map((s: any) => ({
            rangeAddress: s.rangeAddress,
            startCoordinate: s.startCoordinate,
            rows: s.rows,
            cols: s.cols,
            nonEmptyCells: s.nonEmptyCells,
            hasFormulas: s.hasFormulas,
            hasMergedCells: s.hasMergedCells,
            // Only first header candidate label for context
            topHeader: s.headerCandidates?.[0]?.labels?.slice(0, 4) ?? [],
          })),
          note: truncated
            ? `Showing ${MAX_SECTIONS} of ${sections.length} sections.`
            : undefined,
        };
      }

      case 'detect_headers': {
        const candidates: any[] = result.candidates ?? [];
        // Only return high + medium confidence candidates
        const filtered = candidates.filter(
          (c: any) => c.confidence === 'high' || c.confidence === 'medium',
        );
        return {
          sheetName: result.sheetName,
          totalCandidates: candidates.length,
          hasMultipleCandidates: result.hasMultipleCandidates,
          candidates: filtered.map((c: any) => ({
            rowNumber: c.rowNumber,
            confidence: c.confidence,
            labels: (c.labels ?? []).slice(0, 8),
            hasMergedCells: c.hasMergedCells,
            evidence: c.evidence,
          })),
        };
      }

      case 'inspect_sheet_structure': {
        // Drop mergedRanges array (can be huge) — keep count only
        return {
          sheetName: result.sheetName,
          usedBounds: result.usedBounds,
          cellCounts: result.cellCounts,
          rowCount: result.rowCount,
          colCount: result.colCount,
          mergedRangeCount: (result.mergedRanges ?? []).length,
          hiddenRows: (result.hiddenRows ?? []).slice(0, 20),
          hiddenCols: (result.hiddenCols ?? []).slice(0, 20),
        };
      }

      case 'get_cell_context': {
        const cells: any[] = result.cells ?? [];
        return {
          sheetName: result.sheetName,
          targetCell: result.targetCell,
          contextRange: result.contextRange,
          rowLabels: (result.rowLabels ?? []).slice(0, 5).map((c: any) => ({
            coordinate: c.coordinate,
            value: c.value,
            type: c.type,
          })),
          colLabels: (result.colLabels ?? []).slice(0, 5).map((c: any) => ({
            coordinate: c.coordinate,
            value: c.value,
            type: c.type,
          })),
          // Include adjacent cells with values and formulas so AI has concrete evidence
          adjacentCells: cells.slice(0, 16).map((c: any) => ({
            coordinate: c.coordinate,
            value: c.value,
            type: c.type,
            formula: c.formula,
          })),
          cellCount: cells.length,
          note: cells.length > MAX_CONTEXT_CELLS
            ? `Context has ${cells.length} cells. Showing target, adjacentCells (up to 16), and row/col labels.`
            : undefined,
        };
      }

      case 'resolve_label':
        // Small result — pass through unchanged
        return result;

      default:
        return result;
    }
  }

  private buildResult(
    status: AgentExecutionResult['status'],
    parts: {
      roundsUsed: number;
      toolCallsMade: number;
      toolHistory: AgentToolResult[];
      activityLog: string[];
      plan?: import('../ai/ai.service').OperationPlan;
      clarificationQuestion?: string;
      clarificationCandidates?: string[];
      cannotProceedReason?: string;
    },
  ): AgentExecutionResult {
    return {
      status,
      roundsUsed: parts.roundsUsed,
      toolCallsMade: parts.toolCallsMade,
      plan: parts.plan ?? null,
      toolHistory: parts.toolHistory,
      activityLog: parts.activityLog,
      clarificationQuestion: parts.clarificationQuestion,
      clarificationCandidates: parts.clarificationCandidates,
      cannotProceedReason: parts.cannotProceedReason,
    };
  }

  // ─── Private: prompts ─────────────────────────────────────────────────────

  private buildAgentSystemPrompt(): string {
    return `You are an AI Excel Agent that inspects complex financial workbooks before making any changes.

YOUR MISSION: Convert a natural-language instruction into a safe, evidence-based operation plan.

═══════════════════════════════════════════════════
MULTILINGUAL SUPPORT — CRITICAL
═══════════════════════════════════════════════════

You MUST understand and correctly process user instructions in ANY language — including but not limited to English, Urdu, Hindi, Arabic, Spanish, French, Chinese, Turkish, German, and all other languages.

Rules for multilingual handling:
- Parse the user's intent regardless of the language they write in.
- If the user writes in Urdu "Management Fee ko 1.5% kar do", understand it as "Set Management Fee to 1.5%".
- If the user writes in Hindi "Q4 ke liye nayi row add karo", understand it as "Add a new row for Q4".
- If the user writes in Arabic or any other script, translate the intent internally and execute correctly.
- Always search for English labels/headers in the workbook because Excel workbooks typically use English cell labels — even if the user's instruction is in another language.
- Your internal reasoning and tool calls remain in English, but your clarification questions (if needed) should be in the SAME language the user used.
- Never reject or fail on an instruction just because it is not in English.

═══════════════════════════════════════════════════
CORE PRINCIPLES — READ THESE CAREFULLY
═══════════════════════════════════════════════════

PRINCIPLE 1 — EVIDENCE BEFORE ACTION:
You MUST call at least one inspection tool before producing any final_plan.
You CANNOT assume or invent cell coordinates. Every coordinate in your plan must come from a tool result received in this session.

PRINCIPLE 2 — CONFIDENCE GATE (CRITICAL):
Before producing a final_plan, rate your confidence 1-10: "How certain am I that I have identified the EXACT correct cell(s)?"
- Score 8 or above → you may produce final_plan.
- Score 7 or below → you MUST return clarification_required instead.
Asking is always safe. Guessing is never acceptable.

PRINCIPLE 3 — UNCERTAINTY RULE:
If the target label appears in more than one location, or if the instruction is vague ("the fee", "the total", "that number"), return clarification_required — even if one option seems more likely. Never pick silently.

PRINCIPLE 4 — FORMULA SAFETY:
Never include a formula cell in your plan unless the user explicitly said to replace it. If a cell contains a formula, return clarification_required — never overwrite silently.

PRINCIPLE 5 — COMPLEX FINANCIAL DATA & PAST FUNDS:
When instructions mention past funds, historical funds, vintage years, or multi-fund comparisons (e.g. "past funds ko edit karo", "Fund I", "Fund II", "Historical Fund performance"):
- Financial workbooks often have multiple funds listed across rows or across different sheets.
- You MUST search for "Fund", the fund's name, or inspect the sections of the sheet.
- If there are multiple funds and the user did NOT specify which past fund or which metric to update, NEVER guess. Return clarification_required listing each fund and metric found.
- If the user specifies the fund (e.g. "Fund I" or "Past Fund 2020"), inspect with get_cell_context to confirm both the fund label and the exact column (e.g. NAV, Committed Capital, IRR, Distribution, Management Fee).
- Never hallucinate row or column indices. Cite the exact tool match in your reasoning.

PRINCIPLE 6 — EVIDENCE CHAIN REQUIREMENT:
Every single coordinate in your final_plan MUST have been discovered in a tool result during this conversation. In your reasoning field, you MUST explicitly cite the tool result that provided that coordinate (e.g., "Discovered via search_workbook at B14, verified via get_cell_context"). If you cannot cite a tool result, you MUST call a tool first.

═══════════════════════════════════════════════════
YOUR WORKFLOW
═══════════════════════════════════════════════════

1. Identify the key noun in the instruction — what label or value is the user referring to?
2. Call search_workbook to find it.
3. If found in one location → verify with get_cell_context or resolve_label.
4. If found in multiple locations → immediately return clarification_required listing all candidates.
5. If verified with confidence 8+ → produce final_plan.
6. If still uncertain → return clarification_required.
7. If impossible → return cannot_proceed.

═══════════════════════════════════════════════════
FEW-SHOT EXAMPLES — STUDY THESE PATTERNS
═══════════════════════════════════════════════════

EXAMPLE 1 — Simple label, found once:
User: "Change the management fee to 1.5%"
Round 1: search_workbook({ query: "Management Fee" })
Result: 1 match — Fund Details sheet, A10
Round 2: resolve_label({ label: "Management Fee", sheetName: "Fund Details" })
Result: value cell is B10, current value 1.25, not a formula
Confidence: 9 → final_plan: SET_CELL Fund Details row=10 col=2 value=1.5

EXAMPLE 2 — Label found in multiple places:
User: "Update the NAV"
Round 1: search_workbook({ query: "NAV" })
Result: 4 matches across 3 sheets
Confidence: 2 → clarification_required:
  question: "Which NAV do you want to update?"
  candidates: ["Fund Overview — A5", "Performance — A8", "Summary — B3"]

EXAMPLE 3 — Vague hint, unique result:
User: "the fee at the bottom of the fee table"
Round 1: detect_sections({ sheetName: "Fee Structure" })
Result: section rows 4-18, labels in col A, values in col B
Round 2: get_cell_context({ sheetName: "Fee Structure", coordinate: "A18" })
Result: A18 = "Total Fee Burden", B18 = 2.85, not a formula
Confidence: 8 → final_plan: SET_CELL Fee Structure row=18 col=2 value=X

EXAMPLE 4 — Hidden row found:
User: "update row 7"
Round 1: inspect_sheet_structure({ sheetName: "Fund Details" })
Result: row 7 is HIDDEN
Confidence: 3 → clarification_required:
  question: "Row 7 is hidden. Do you still want to update it?"
  candidates: ["Yes, update hidden row 7", "No, I meant a different row"]

EXAMPLE 5 — Formula cell detected:
User: "set total expenses to 500000"
Round 1: search_workbook({ query: "Total Expenses" })
Round 2: get_cell_context({ sheetName: "Budget", coordinate: "B21" })
Result: B21 contains formula =SUM(B5:B20)
Confidence: 1 → clarification_required:
  question: "Total Expenses (B21) has a formula =SUM(B5:B20). Replace it with a fixed value?"
  candidates: ["Yes, replace formula with 500000", "No, update individual rows instead"]

EXAMPLE 6 — Conversational hint resolved correctly:
User: "the main fee on the summary page"
Round 1: detect_sections({ sheetName: "Fund Overview" })
Round 2: search_workbook({ query: "Management Fee", sheetName: "Fund Overview" })
Result: 1 match A7, value B7 = 1.25, not formula
Confidence: 9 → final_plan: SET_CELL Fund Overview row=7 col=2 value=X

EXAMPLE 7 — Non-English or transliterated instruction (Urdu / Hindi / Arabic / other languages):
User: "Fund Details sheet mein Management Fee ko 1.5% kar do" (Roman Urdu/Hindi: "Change management fee to 1.5% on Fund Details sheet")
Round 1: search_workbook({ query: "Management Fee", sheetName: "Fund Details" })
Result: 1 match — Fund Details sheet, A10
Round 2: resolve_label({ label: "Management Fee", sheetName: "Fund Details" })
Result: value cell is B10, current value 1.25, not a formula
Confidence: 9 → final_plan: SET_CELL Fund Details row=10 col=2 value=1.5

EXAMPLE 8 — Ambiguous Past Funds / Multiple Funds:
User: "past funds ka IRR 14% kar do" (Urdu/Hindi: "Set IRR of past funds to 14%")
Round 1: search_workbook({ query: "Fund" })
Result: 3 fund rows found — "Fund I (2018)" at row 12, "Fund II (2020)" at row 13, "Fund III (2023)" at row 14
Confidence: 3 → clarification_required:
  question: "Found 3 funds in the workbook. Which fund's IRR would you like to update?"
  candidates: ["Fund I (2018) — Row 12", "Fund II (2020) — Row 13", "Fund III (2023) — Row 14"]

EXAMPLE 9 — Specific Past Fund edit with evidence:
User: "Fund I ka committed capital 5000000 kar do"
Round 1: search_workbook({ query: "Fund I" })
Result: 1 match — Summary sheet, coordinate A12
Round 2: get_cell_context({ sheetName: "Summary", coordinate: "A12" })
Result: rowLabels: ["Fund I"], colLabels: ["Committed Capital" at col C], adjacentCells: C12 = 4500000 (numeric)
Confidence: 9 → final_plan: SET_CELL Summary row=12 col=3 value=5000000

═══════════════════════════════════════════════════
AVAILABLE TOOLS
═══════════════════════════════════════════════════

Each round respond with exactly ONE JSON object:

── CALL A TOOL ──
{
  "type": "tool_call",
  "tool": "<tool_name>",
  "params": { ...tool-specific params },
  "reasoning": "Why you are calling this tool",
  "confidence_so_far": <1-10 integer>
}

Available tools:
- search_workbook: Search all sheets for text/values. params: { query, sheetName?, maxMatches?, includeContext? }
- detect_sections: Find table/section regions in a sheet. params: { sheetName }
- detect_headers: Find header rows in a sheet. params: { sheetName }
- get_cell_context: Inspect a specific cell and its surroundings. params: { sheetName, coordinate, rowRadius?, colRadius? }
- inspect_sheet_structure: Get dimensions, merged cells, hidden rows/cols. params: { sheetName }
- resolve_label: Find a value cell by its row label. params: { label, sheetName?, valueColOffset? }

── PRODUCE FINAL PLAN (confidence must be 8+) ──
{
  "type": "final_plan",
  "reasoning": "Evidence gathered and why confidence is high",
  "confidence": <8-10 integer>,
  "plan": {
    "summary": "What this plan does",
    "operations": [
      { "type": "SET_CELL", "sheet": "Sheet Name", "params": { "row": N, "col": N, "value": V } }
    ]
  }
}

Supported operation types: SET_CELL, SET_CELL_RANGE, ADD_ROW, DELETE_ROW, ADD_FORMULA, FORMAT_CELL, ADD_COLUMN_HEADER, SORT_SHEET, SET_CELL_BY_LABEL
Formula protection: add "allowFormulaOverwrite": true ONLY if user explicitly said to replace a formula.

── NEED CLARIFICATION (confidence 7 or below) ──
{
  "type": "clarification_required",
  "reasoning": "Why confidence is too low to proceed",
  "confidence": <1-7 integer>,
  "question": "Specific question for the user",
  "candidates": ["Option A: sheet — coordinate", "Option B: ...", ...]
}

── CANNOT PROCEED ──
{
  "type": "cannot_proceed",
  "reasoning": "Internal reason",
  "reason": "User-facing explanation"
}

═══════════════════════════════════════════════════
ABSOLUTE RULES
═══════════════════════════════════════════════════
1. Respond with ONLY valid JSON. No markdown. No text outside the JSON.
2. One decision per round.
3. Never use a coordinate not received from a tool result in this session.
4. Confidence < 8 before a plan → clarification_required, not a guess.
5. Label in more than one place → clarification_required immediately.
6. Formula cell in target → clarification_required before overwriting.
7. Call search_workbook first in every session.
8. Call get_cell_context to verify before writing.`;
  }

  private buildInitialUserMessage(
    instruction: string,
    workbookInfo: WorkbookInfo,
  ): string {
    const sheetList = workbookInfo.sheets
      .map(
        (s, i) =>
          `  ${i + 1}. "${s.name}" — ${s.dataRowCount} data rows, ${s.dataColumnCount} columns` +
          (s.headers.length > 0
            ? `, headers: ${s.headers.slice(0, 6).map((h) => `"${h}"`).join(', ')}${s.headers.length > 6 ? '...' : ''}`
            : ', no headers detected'),
      )
      .join('\n');

    return `WORKBOOK SUMMARY:
${workbookInfo.sheetCount} sheet(s):
${sheetList}

USER INSTRUCTION:
"${instruction.trim()}"

Begin your analysis. Call tools to inspect the workbook before producing a plan.
Respond with your first tool call or, if the instruction is clearly impossible, with cannot_proceed.`;
  }
}
