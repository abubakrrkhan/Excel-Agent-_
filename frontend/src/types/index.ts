// ─── Backend entity types ─────────────────────────────────────────────────────

export type FileStatus = 'PENDING' | 'PROCESSING' | 'DONE' | 'ERROR';

export type AgentJobStatus =
  | 'queued'
  | 'processing'
  | 'completed'
  | 'clarification_needed'
  | 'cannot_proceed'
  | 'failed';

export interface HealthResponse {
  status: string;
  timestamp: string;
  environment: string;
  service: string;
  version: string;
}

export interface ApiError {
  message: string;
  statusCode?: number;
}

// ─── Job polling response ─────────────────────────────────────────────────────

export interface JobStatusResponse {
  status: AgentJobStatus;
  jobId: string;
  // Completed
  outputStoredName?: string;
  downloadUrl?: string;
  operationSummary?: string;
  operationsApplied?: number;
  writeOpsApplied?: number;
  // Agent activity
  agentActivityLog?: string[];
  agentRoundsUsed?: number;
  agentToolCallsMade?: number;
  // Clarification
  clarificationQuestion?: string;
  clarificationCandidates?: string[];
  // Error
  cannotProceedReason?: string;
  error?: string;
  finishedAt?: string;
  progress?: number;
}

// ─── UI state types ───────────────────────────────────────────────────────────

export type AppStage =
  | 'idle'
  | 'ready'
  | 'submitting'
  | 'queued'
  | 'processing'
  | 'clarification'
  | 'completed'
  | 'failed';

export interface SelectedFile {
  file: File;
  name: string;
  sizeBytes: number;
}

export interface FormErrors {
  file?: string;
  instruction?: string;
}

// Legacy alias kept for backwards compat
export interface ProcessResult {
  outputFileId: string;
  outputStoredName: string;
  downloadUrl: string;
  originalName: string;
  operationSummary: string;
  operationsApplied: number;
}
