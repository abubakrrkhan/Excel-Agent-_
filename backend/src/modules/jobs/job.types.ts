/**
 * job.types.ts
 *
 * Phase 6 — Async job queue types.
 * Defines the data shapes for BullMQ job payloads and results.
 */

import { OperationPlan } from '../ai/ai.service';

export const EXCEL_AGENT_QUEUE = 'excel-agent';
export const EXCEL_AGENT_JOB  = 'run-agent';

// ─── Job payload (what's sent INTO the queue) ─────────────────────────────────

export interface AgentJobPayload {
  jobId: string;
  /**
   * Stored filename on disk (UUID + .xlsx).
   * The processor reads the file via FilesService.readFileFromDisk().
   * We NEVER embed the raw buffer or base64 in the queue payload —
   * a 10 MB file becomes ~13 MB of base64, which quickly exhausts Redis memory.
   */
  storedFileName: string;
  /** Original filename for output naming */
  originalName: string;
  /** User's natural-language instruction */
  instruction: string;
  /**
   * Optional clarification answer from a previous round.
   * When present, the processor prepends it to the instruction.
   */
  clarificationAnswer?: string;
}

// ─── Job result (what comes OUT of the processor) ────────────────────────────

export type JobStatus =
  | 'queued'
  | 'processing'
  | 'completed'
  | 'clarification_needed'
  | 'cannot_proceed'
  | 'failed';

export interface AgentJobResult {
  status: JobStatus;
  jobId: string;
  /** Only present when status = 'completed' */
  outputStoredName?: string;
  downloadUrl?: string;
  operationSummary?: string;
  operationsApplied?: number;
  writeOpsApplied?: number;
  plan?: OperationPlan | null;
  /** Agent activity steps */
  agentActivityLog?: string[];
  agentRoundsUsed?: number;
  agentToolCallsMade?: number;
  /** Only when status = 'clarification_needed' */
  clarificationQuestion?: string;
  clarificationCandidates?: string[];
  /** Only when status = 'cannot_proceed' */
  cannotProceedReason?: string;
  /** Only when status = 'failed' */
  error?: string;
  /** ISO timestamp when job completed/failed */
  finishedAt?: string;
}
