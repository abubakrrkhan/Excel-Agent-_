/**
 * agent.types.ts
 *
 * Phase 3 Step 1 — AI Agent Core type definitions.
 *
 * The agent uses a tool-calling loop:
 *   1. AI receives workbook context + available tools
 *   2. AI calls a tool (inspection or label resolution)
 *   3. Tool result is returned to AI
 *   4. AI reasons over accumulated evidence
 *   5. Repeat until AI produces a final OperationPlan or requests clarification
 *
 * Key principles:
 *   - AI never touches the Excel file directly
 *   - Tools are the only way to inspect the workbook
 *   - Every write operation requires prior inspection evidence
 *   - Ambiguous targets surface as ClarificationRequired, never silently chosen
 */

import { OperationPlan } from '../ai/ai.service';

// ─── Tool definitions (what the AI can call) ─────────────────────────────────

/**
 * Names of tools the agent can invoke.
 * These map 1:1 to existing inspection services.
 */
export type AgentToolName =
  | 'inspect_sheet_structure'
  | 'detect_headers'
  | 'detect_sections'
  | 'get_cell_context'
  | 'search_workbook'
  | 'resolve_label';

/**
 * A tool call requested by the AI in a given round.
 */
export interface AgentToolCall {
  /** Which tool to invoke */
  tool: AgentToolName;
  /** Tool-specific parameters */
  params: Record<string, any>;
  /** AI's stated reason for calling this tool (for logging/transparency) */
  reason?: string;
}

/**
 * Result of executing a single tool call.
 */
export interface AgentToolResult {
  tool: AgentToolName;
  params: Record<string, any>;
  /** Structured data returned by the tool */
  result: Record<string, any>;
  /** Whether the tool call succeeded */
  success: boolean;
  /** Error message if success=false */
  error?: string;
}

// ─── Agent conversation message ───────────────────────────────────────────────

export type AgentMessageRole = 'system' | 'user' | 'assistant' | 'tool';

export interface AgentMessage {
  role: AgentMessageRole;
  content: string;
}

// ─── AI response in each round ───────────────────────────────────────────────

/**
 * What the AI returns in each reasoning round.
 *
 * The AI produces ONE of:
 *   1. tool_call  — invoke a tool to gather more evidence
 *   2. final_plan — enough evidence gathered, produce the operation plan
 *   3. clarification_required — ambiguous target, ask the user
 *   4. cannot_proceed — instruction cannot be fulfilled safely
 */
export type AgentRoundType =
  | 'tool_call'
  | 'final_plan'
  | 'clarification_required'
  | 'cannot_proceed';

export interface AgentRoundResponseBase {
  type: AgentRoundType;
  reasoning: string; // AI's internal reasoning (logged, not shown to user)
}

export interface AgentRoundToolCall extends AgentRoundResponseBase {
  type: 'tool_call';
  /** Which tool to invoke */
  tool: AgentToolName;
  /** Tool-specific parameters */
  params: Record<string, any>;
}

export interface AgentRoundFinalPlan extends AgentRoundResponseBase {
  type: 'final_plan';
  plan: OperationPlan;
}

export interface AgentRoundClarification extends AgentRoundResponseBase {
  type: 'clarification_required';
  question: string;         // User-facing question
  candidates: string[];     // Candidate descriptions to help user choose
}

export interface AgentRoundCannotProceed extends AgentRoundResponseBase {
  type: 'cannot_proceed';
  reason: string;           // User-facing explanation
}

export type AgentRoundResponse =
  | AgentRoundToolCall
  | AgentRoundFinalPlan
  | AgentRoundClarification
  | AgentRoundCannotProceed;

// ─── Agent execution result ───────────────────────────────────────────────────

export type AgentResultStatus =
  | 'completed'             // plan produced and executed successfully
  | 'clarification_needed'  // ambiguous — user must clarify before proceeding
  | 'cannot_proceed'        // agent determined the instruction is not safe/possible
  | 'max_rounds_reached';   // safety limit hit before a conclusion was reached

export interface AgentExecutionResult {
  status: AgentResultStatus;
  /** Number of AI reasoning rounds used */
  roundsUsed: number;
  /** Number of tool calls made */
  toolCallsMade: number;
  /** The final operation plan (only when status = 'completed') */
  plan: OperationPlan | null;
  /** Tool call history — all inspections performed */
  toolHistory: AgentToolResult[];
  /** If clarification_needed: the question to show the user */
  clarificationQuestion?: string;
  /** If clarification_needed: candidate options to present */
  clarificationCandidates?: string[];
  /** If cannot_proceed: reason to show the user */
  cannotProceedReason?: string;
  /** Human-readable agent activity log (for future ChatGPT-style UI) */
  activityLog: string[];
}

// ─── Agent configuration ──────────────────────────────────────────────────────

export interface AgentConfig {
  /**
   * Maximum number of AI reasoning rounds.
   * Safety limit — prevents infinite loops.
   * Default: 8
   */
  maxRounds: number;
  /**
   * Maximum number of tool calls per session.
   * Prevents runaway inspection.
   * Default: 12
   */
  maxToolCalls: number;
}

export const DEFAULT_AGENT_CONFIG: AgentConfig = {
  maxRounds: 8,
  maxToolCalls: 12,
};

/**
 * Config for complex workbooks — more sheets, more data.
 * Used automatically when a workbook has 3+ sheets OR 500+ data rows.
 * First INSPECTION_ONLY_ROUNDS rounds are read-only — no writes permitted.
 */
export const COMPLEX_AGENT_CONFIG: AgentConfig = {
  maxRounds: 14,
  maxToolCalls: 20,
};

/** How many rounds must be pure inspection before writes are allowed */
export const INSPECTION_ONLY_ROUNDS = 3;

/**
 * Detect whether a workbook qualifies as "complex" and needs more rounds.
 */
export function isComplexWorkbook(workbookInfo: { sheetCount: number; sheets: Array<{ dataRowCount: number }> }): boolean {
  const totalRows = workbookInfo.sheets.reduce((sum, s) => sum + (s.dataRowCount || 0), 0);
  return workbookInfo.sheetCount >= 3 || totalRows >= 500;
}
