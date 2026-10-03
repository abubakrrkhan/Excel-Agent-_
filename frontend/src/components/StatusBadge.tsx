'use client';

import type { HealthResponse } from '@/types';

interface StatusBadgeProps {
  health: HealthResponse | null;
  error: string | null;
}

export default function StatusBadge({ health, error }: StatusBadgeProps) {
  const s = getStatus(health, error);

  return (
    <div style={styles.wrap} role="status" aria-label={`Backend status: ${s.label}`}>
      <span style={{ ...styles.dot, background: s.color }} aria-hidden="true" />
      <span style={styles.text}>{s.label}</span>
    </div>
  );
}

function getStatus(health: HealthResponse | null, error: string | null) {
  if (error)  return { color: 'var(--error)',   label: 'Backend offline' };
  if (health) return { color: 'var(--success)', label: `API connected — v${health.version}` };
  return       { color: 'var(--text-subtle)',   label: 'Connecting…' };
}

const styles: Record<string, React.CSSProperties> = {
  wrap: {
    display: 'flex',
    alignItems: 'center',
    gap: '0.5rem',
    fontSize: '0.82rem',
    color: 'var(--text-muted)',
  },
  dot: {
    width: 8,
    height: 8,
    borderRadius: '50%',
    flexShrink: 0,
  },
  text: {
    fontVariantNumeric: 'tabular-nums',
  },
};
