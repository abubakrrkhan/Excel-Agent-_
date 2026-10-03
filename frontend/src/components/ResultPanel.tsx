'use client';

import type { ProcessResult, AppStage } from '@/types';

interface ResultPanelProps {
  state: AppStage;
  result: ProcessResult | null;
  error: string | null;
}

export default function ResultPanel({ state, result, error }: ResultPanelProps) {
  // Only render when there is something to show
  if (state === 'idle' || state === 'ready') return null;

  return (
    <section aria-labelledby="result-label" aria-live="polite">
      <p id="result-label" style={styles.sectionLabel}>
        Result
      </p>

      {/* ── Processing ──────────────────────────────────────────────────── */}
      {state === 'processing' && (
        <div style={styles.card} role="status">
          <div style={styles.processingRow}>
            <span style={styles.spinner} aria-hidden="true" />
            <div>
              <p style={styles.processingTitle}>Processing your file…</p>
              <p style={styles.processingHint}>
                The AI is reading your workbook and generating the operation plan.
                This usually takes a few seconds.
              </p>
            </div>
          </div>
        </div>
      )}

      {/* ── Error ───────────────────────────────────────────────────────── */}
      {state === 'failed' && error && (
        <div style={{ ...styles.card, ...styles.cardError }} role="alert">
          <div style={styles.statusRow}>
            <span style={styles.iconError} aria-hidden="true">✕</span>
            <p style={styles.errorTitle}>Processing failed</p>
          </div>
          <p style={styles.errorMessage}>{error}</p>
        </div>
      )}

      {/* ── Success ─────────────────────────────────────────────────────── */}
      {state === 'completed' && result && (
        <div style={{ ...styles.card, ...styles.cardSuccess }}>
          {/* Status row */}
          <div style={styles.statusRow}>
            <span style={styles.iconSuccess} aria-hidden="true">✓</span>
            <p style={styles.successTitle}>File processed successfully</p>
          </div>

          {/* AI summary */}
          <div style={styles.summaryBox}>
            <p style={styles.summaryLabel}>AI applied:</p>
            <p style={styles.summaryText}>{result.operationSummary}</p>
          </div>

          {/* Stats row */}
          <div style={styles.statsRow}>
            <StatChip
              label="Operations applied"
              value={String(result.operationsApplied)}
            />
            <StatChip label="Original file" value={result.originalName} />
          </div>

          {/* Download button — wired in Day 4 Part 2 */}
          <a
            href={result.downloadUrl}
            download
            style={styles.downloadBtn}
            aria-label={`Download processed file: ${result.originalName}`}
          >
            ⬇ Download processed file
          </a>
        </div>
      )}
    </section>
  );
}

// ─── Sub-component ────────────────────────────────────────────────────────────

function StatChip({ label, value }: { label: string; value: string }) {
  return (
    <div style={styles.chip}>
      <span style={styles.chipLabel}>{label}</span>
      <span style={styles.chipValue}>{value}</span>
    </div>
  );
}

// ─── Styles ───────────────────────────────────────────────────────────────────

const styles: Record<string, React.CSSProperties> = {
  sectionLabel: {
    fontWeight: 600,
    fontSize: '0.9rem',
    color: 'var(--text)',
    marginBottom: '0.6rem',
  },
  card: {
    background: 'var(--surface)',
    border: '1px solid var(--border)',
    borderRadius: 'var(--radius-lg)',
    padding: '1.25rem',
    display: 'flex',
    flexDirection: 'column',
    gap: '1rem',
  },
  cardError: {
    borderColor: 'var(--error-border)',
    background: 'var(--error-bg)',
  },
  cardSuccess: {
    borderColor: 'var(--success-border)',
    background: 'var(--success-bg)',
  },
  // Processing
  processingRow: {
    display: 'flex',
    alignItems: 'flex-start',
    gap: '1rem',
  },
  spinner: {
    display: 'inline-block',
    width: 22,
    height: 22,
    border: '2px solid var(--border)',
    borderTopColor: 'var(--accent)',
    borderRadius: '50%',
    flexShrink: 0,
    marginTop: 2,
    animation: 'spin 0.8s linear infinite',
  },
  processingTitle: {
    fontWeight: 600,
    fontSize: '0.9rem',
    color: 'var(--text)',
  },
  processingHint: {
    fontSize: '0.82rem',
    color: 'var(--text-muted)',
    marginTop: '0.25rem',
  },
  // Status row (success / error)
  statusRow: {
    display: 'flex',
    alignItems: 'center',
    gap: '0.6rem',
  },
  iconSuccess: {
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    width: 22,
    height: 22,
    borderRadius: '50%',
    background: 'var(--success)',
    color: '#fff',
    fontSize: '0.75rem',
    fontWeight: 700,
    flexShrink: 0,
  },
  iconError: {
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    width: 22,
    height: 22,
    borderRadius: '50%',
    background: 'var(--error)',
    color: '#fff',
    fontSize: '0.75rem',
    fontWeight: 700,
    flexShrink: 0,
  },
  successTitle: {
    fontWeight: 600,
    fontSize: '0.9rem',
    color: 'var(--success)',
  },
  errorTitle: {
    fontWeight: 600,
    fontSize: '0.9rem',
    color: 'var(--error)',
  },
  errorMessage: {
    fontSize: '0.85rem',
    color: 'var(--text-muted)',
    lineHeight: 1.5,
  },
  // Summary box
  summaryBox: {
    background: 'rgba(0,0,0,0.15)',
    borderRadius: 'var(--radius)',
    padding: '0.75rem 1rem',
  },
  summaryLabel: {
    fontSize: '0.78rem',
    fontWeight: 600,
    color: 'var(--text-muted)',
    textTransform: 'uppercase',
    letterSpacing: '0.05em',
    marginBottom: '0.3rem',
  },
  summaryText: {
    fontSize: '0.88rem',
    color: 'var(--text)',
    lineHeight: 1.5,
  },
  // Stats
  statsRow: {
    display: 'flex',
    flexWrap: 'wrap' as const,
    gap: '0.5rem',
  },
  chip: {
    background: 'rgba(0,0,0,0.2)',
    borderRadius: 'var(--radius)',
    padding: '0.4rem 0.75rem',
    display: 'flex',
    flexDirection: 'column' as const,
    gap: '0.1rem',
    minWidth: 0,
    maxWidth: '100%',
  },
  chipLabel: {
    fontSize: '0.7rem',
    fontWeight: 600,
    color: 'var(--text-muted)',
    textTransform: 'uppercase' as const,
    letterSpacing: '0.04em',
  },
  chipValue: {
    fontSize: '0.85rem',
    color: 'var(--text)',
    fontWeight: 500,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap' as const,
  },
  // Download button
  downloadBtn: {
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    background: 'var(--success)',
    color: '#fff',
    border: 'none',
    borderRadius: 'var(--radius)',
    padding: '0.7rem 1.5rem',
    fontWeight: 600,
    fontSize: '0.9rem',
    textDecoration: 'none',
    transition: 'opacity 0.15s',
    alignSelf: 'flex-start' as const,
  },
};
