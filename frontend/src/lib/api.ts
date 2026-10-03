import axios from 'axios';
import type { HealthResponse, JobStatusResponse } from '@/types';

const API_BASE_URL = process.env.NEXT_PUBLIC_API_URL || 'http://localhost:3001';

export const apiClient = axios.create({
  baseURL: `${API_BASE_URL}/api`,
  timeout: 30_000,
  headers: { 'Content-Type': 'application/json' },
});

apiClient.interceptors.response.use(
  (r) => r,
  (e) => {
    const message = e.response?.data?.message || e.message || 'Something went wrong';
    return Promise.reject(new Error(message));
  },
);

// ─── Health ───────────────────────────────────────────────────────────────────

export const healthCheck = () => apiClient.get<HealthResponse>('/health');

// ─── Async Job Submit ─────────────────────────────────────────────────────────

export async function submitJob(
  file: File,
  instruction: string,
): Promise<{ jobId: string; pollUrl: string }> {
  const form = new FormData();
  form.append('file', file, file.name);
  form.append('instruction', instruction.trim());

  const res = await apiClient.post<{ jobId: string; pollUrl: string; status: string }>(
    '/jobs/submit',
    form,
    { headers: { 'Content-Type': undefined } },
  );
  return { jobId: res.data.jobId, pollUrl: res.data.pollUrl };
}

// ─── Poll Job Status ──────────────────────────────────────────────────────────

export async function pollJobStatus(jobId: string): Promise<JobStatusResponse> {
  const res = await apiClient.get<JobStatusResponse>(`/jobs/${jobId}`);
  return res.data;
}

// ─── Clarification re-submit ──────────────────────────────────────────────────

/**
 * Re-submit after user answers a clarification question.
 * The backend combines instruction + clarificationAnswer into the agent prompt.
 */
export async function submitClarification(
  file: File,
  instruction: string,
  clarificationAnswer: string,
): Promise<{ jobId: string }> {
  const form = new FormData();
  form.append('file', file, file.name);
  form.append('instruction', instruction.trim());
  form.append('clarificationAnswer', clarificationAnswer.trim());

  const res = await apiClient.post<{ jobId: string; status: string; pollUrl: string }>(
    '/jobs/submit',
    form,
    { headers: { 'Content-Type': undefined } },
  );
  return { jobId: res.data.jobId };
}
