'use client';

import { useState, useRef, useCallback, useEffect } from 'react';
import { submitJob, pollJobStatus, submitClarification } from '@/lib/api';
import type { AppStage, JobStatusResponse, SelectedFile } from '@/types';

// ─── Types ────────────────────────────────────────────────────────────────────

interface HistoryItem {
  id: string;
  label: string;
  timestamp: number;
  status: 'completed' | 'failed' | 'clarification_needed' | 'cannot_proceed';
}

interface Theme {
  bg: string;
  sidebar: string;
  card: string;
  input: string;
  border: string;
  borderSubtle: string;
  text: string;
  textSecondary: string;
  textMuted: string;
  logBg: string;
  logText: string;
  btnHover: string;
}

const DARK: Theme = {
  bg: '#0a0a0a',
  sidebar: '#0f0f0f',
  card: '#1a1a1a',
  input: '#1a1a1a',
  border: 'rgba(255,255,255,0.1)',
  borderSubtle: 'rgba(255,255,255,0.06)',
  text: '#ececec',
  textSecondary: '#8e8ea0',
  textMuted: '#4d4d4d',
  logBg: '#111',
  logText: '#c9d1d9',
  btnHover: 'rgba(255,255,255,0.05)',
};

const LIGHT: Theme = {
  bg: '#f5f5f5',
  sidebar: '#ffffff',
  card: '#ffffff',
  input: '#ffffff',
  border: 'rgba(0,0,0,0.1)',
  borderSubtle: 'rgba(0,0,0,0.07)',
  text: '#111827',
  textSecondary: '#6b7280',
  textMuted: '#9ca3af',
  logBg: '#f3f4f6',
  logText: '#374151',
  btnHover: 'rgba(0,0,0,0.04)',
};

// ─── Helpers ──────────────────────────────────────────────────────────────────

function formatBytes(b: number) {
  if (b < 1024) return `${b} B`;
  if (b < 1048576) return `${(b / 1024).toFixed(1)} KB`;
  return `${(b / 1048576).toFixed(1)} MB`;
}

function truncate(s: string, n = 28) {
  return s.length > n ? s.slice(0, n) + '…' : s;
}

// ─── Icons ────────────────────────────────────────────────────────────────────

const Icon = {
  grid: (
    <svg width="18" height="18" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.8}>
      <rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/>
      <rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/>
    </svg>
  ),
  newChat: (
    <svg width="16" height="16" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M12 4v16m8-8H4"/>
    </svg>
  ),
  history: (
    <svg width="14" height="14" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.8}>
      <circle cx="12" cy="12" r="9"/><path strokeLinecap="round" d="M12 7v5l3 3"/>
    </svg>
  ),
  file: (
    <svg width="16" height="16" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.8}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/>
    </svg>
  ),
  plus: (
    <svg width="18" height="18" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M12 4v16m8-8H4"/>
    </svg>
  ),
  send: (
    <svg width="15" height="15" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.5}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M4.5 10.5L12 3m0 0l7.5 7.5M12 3v18"/>
    </svg>
  ),
  check: (
    <svg width="14" height="14" fill="none" viewBox="0 0 24 24" stroke="#10a37f" strokeWidth={2.5}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M4.5 12.75l6 6 9-13.5"/>
    </svg>
  ),
  x: (
    <svg width="14" height="14" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M6 18L18 6M6 6l12 12"/>
    </svg>
  ),
  warn: (
    <svg width="14" height="14" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M12 9v3.75m-9.303 3.376c-.866 1.5.217 3.374 1.948 3.374h14.71c1.73 0 2.813-1.874 1.948-3.374L13.949 3.378c-.866-1.5-3.032-1.5-3.898 0L2.697 16.126zM12 15.75h.007v.008H12v-.008z"/>
    </svg>
  ),
  moon: (
    <svg width="16" height="16" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M21.752 15.002A9.718 9.718 0 0118 15.75 9.75 9.75 0 018.25 6c0-1.33.266-2.597.748-3.752A9.753 9.753 0 003 12a9.75 9.75 0 009.75 9.75 9.753 9.753 0 008.502-4.998z"/>
    </svg>
  ),
  sun: (
    <svg width="16" height="16" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M12 3v2.25m6.364.386l-1.591 1.591M21 12h-2.25m-.386 6.364l-1.591-1.591M12 18.75V21m-4.773-4.227l-1.591 1.591M5.25 12H3m4.227-4.773L5.636 5.636M15.75 12a3.75 3.75 0 11-7.5 0 3.75 3.75 0 017.5 0z"/>
    </svg>
  ),
};

// ─── Activity log ─────────────────────────────────────────────────────────────

function ActivityLog({ lines, running, t }: { lines: string[]; running: boolean; t: Theme }) {
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => { if (ref.current) ref.current.scrollTop = ref.current.scrollHeight; }, [lines]);
  return (
    <div ref={ref} style={{ background: t.logBg, borderRadius: 10, padding: '12px 14px', maxHeight: 200, overflowY: 'auto', fontFamily: 'ui-monospace, monospace', fontSize: 11.5, lineHeight: 1.9, border: `1px solid ${t.borderSubtle}` }}>
      {lines.length === 0
        ? <span style={{ color: t.textMuted }}>Waiting for agent...</span>
        : lines.map((l, i) => (
          <div key={i}>
            <span style={{ color: t.textMuted, marginRight: 10, userSelect: 'none' }}>{String(i + 1).padStart(2, '0')}</span>
            <span style={{ color: t.logText }}>{l}</span>
          </div>
        ))
      }
      {running && lines.length > 0 && <span className="blink" style={{ color: t.textMuted }}>|</span>}
    </div>
  );
}

// ─── Clarification ────────────────────────────────────────────────────────────

function ClarificationPanel({ question, candidates, onAnswer, t }: { question: string; candidates: string[]; onAnswer: (a: string) => void; t: Theme }) {
  const [custom, setCustom] = useState('');
  const [selected, setSelected] = useState<string | null>(null);
  return (
    <div style={{ background: t.card, border: `1px solid ${t.border}`, borderRadius: 12, padding: '18px 20px' }}>
      <p style={{ fontSize: 13, fontWeight: 600, color: t.text, marginBottom: 4 }}>The agent needs clarification</p>
      <p style={{ fontSize: 13, color: t.textSecondary, marginBottom: 14 }}>{question}</p>
      {candidates.length > 0 && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 7, marginBottom: 14 }}>
          {candidates.map((c, i) => (
            <button key={i} onClick={() => { setSelected(c); setCustom(''); }}
              style={{ textAlign: 'left', padding: '9px 14px', borderRadius: 8, border: selected === c ? '1px solid #10a37f' : `1px solid ${t.border}`, background: selected === c ? 'rgba(16,163,127,0.08)' : t.bg, color: selected === c ? '#10a37f' : t.text, fontSize: 13, cursor: 'pointer', transition: 'all 0.12s' }}
            >{c}</button>
          ))}
        </div>
      )}
      <div style={{ display: 'flex', gap: 8 }}>
        <input value={custom} onChange={e => { setCustom(e.target.value); setSelected(null); }} placeholder="Or type a custom answer..."
          style={{ flex: 1, background: t.bg, border: `1px solid ${t.border}`, borderRadius: 8, padding: '9px 12px', fontSize: 13, color: t.text }}
        />
        <button onClick={() => onAnswer(custom.trim() || selected || '')} disabled={!custom.trim() && !selected}
          style={{ padding: '9px 16px', background: '#10a37f', color: '#fff', borderRadius: 8, fontSize: 13, fontWeight: 600, border: 'none', cursor: 'pointer', opacity: (!custom.trim() && !selected) ? 0.4 : 1 }}
        >Confirm</button>
      </div>
    </div>
  );
}

// ─── Sidebar ─────────────────────────────────────────────────────────────────

function Sidebar({ history, activeId, onNew, isDark, onToggleTheme, t }: {
  history: HistoryItem[];
  activeId: string | null;
  onNew: () => void;
  isDark: boolean;
  onToggleTheme: () => void;
  t: Theme;
}) {
  const grouped = {
    today: history.filter(h => Date.now() - h.timestamp < 86400000),
    earlier: history.filter(h => Date.now() - h.timestamp >= 86400000),
  };

  return (
    <aside style={{ width: 260, flexShrink: 0, background: t.sidebar, display: 'flex', flexDirection: 'column', height: '100vh', borderRight: `1px solid ${t.borderSubtle}` }}>
      {/* Logo */}
      <div style={{ padding: '16px 14px 12px', display: 'flex', alignItems: 'center', gap: 10 }}>
        <span style={{ color: '#10a37f' }}>{Icon.grid}</span>
        <span style={{ fontSize: 15, fontWeight: 600, color: t.text }}>Excel AI Agent</span>
      </div>

      {/* New task */}
      <div style={{ padding: '4px 10px 12px' }}>
        <button onClick={onNew}
          style={{ width: '100%', display: 'flex', alignItems: 'center', gap: 10, padding: '9px 12px', borderRadius: 8, background: isDark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.04)', border: `1px solid ${t.border}`, color: t.text, fontSize: 13, cursor: 'pointer', transition: 'background 0.15s' }}
          onMouseEnter={e => (e.currentTarget.style.background = t.btnHover)}
          onMouseLeave={e => (e.currentTarget.style.background = isDark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.04)')}
        >
          <span style={{ opacity: 0.7 }}>{Icon.newChat}</span>
          New task
        </button>
      </div>

      {/* History */}
      <div style={{ flex: 1, overflowY: 'auto', padding: '0 10px' }}>
        {history.length === 0
          ? <p style={{ fontSize: 12, color: t.textMuted, padding: '8px 4px' }}>No previous tasks</p>
          : (
            <>
              {grouped.today.length > 0 && (
                <>
                  <p style={{ fontSize: 11, color: t.textMuted, padding: '6px 4px 4px', textTransform: 'uppercase', letterSpacing: '0.06em', fontWeight: 500 }}>Today</p>
                  {grouped.today.map(h => <HistoryRow key={h.id} item={h} active={h.id === activeId} t={t} />)}
                </>
              )}
              {grouped.earlier.length > 0 && (
                <>
                  <p style={{ fontSize: 11, color: t.textMuted, padding: '10px 4px 4px', textTransform: 'uppercase', letterSpacing: '0.06em', fontWeight: 500 }}>Earlier</p>
                  {grouped.earlier.map(h => <HistoryRow key={h.id} item={h} active={h.id === activeId} t={t} />)}
                </>
              )}
            </>
          )
        }
      </div>

      {/* Footer — user + theme toggle */}
      <div style={{ padding: '12px 14px', borderTop: `1px solid ${t.borderSubtle}` }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <div style={{ width: 28, height: 28, borderRadius: '50%', background: '#10a37f', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 12, fontWeight: 700, color: '#fff', flexShrink: 0 }}>A</div>
          <div style={{ flex: 1 }}>
            <p style={{ fontSize: 13, color: t.text, fontWeight: 500 }}>Abubakr Khan</p>
            <p style={{ fontSize: 11, color: t.textMuted }}>Excel AI Agent</p>
          </div>
          {/* Theme toggle button */}
          <button
            onClick={onToggleTheme}
            title={isDark ? 'Switch to light mode' : 'Switch to dark mode'}
            style={{ width: 32, height: 32, borderRadius: 8, background: isDark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.06)', border: `1px solid ${t.border}`, display: 'flex', alignItems: 'center', justifyContent: 'center', cursor: 'pointer', color: t.textSecondary, flexShrink: 0, transition: 'all 0.2s' }}
            onMouseEnter={e => { e.currentTarget.style.color = t.text; e.currentTarget.style.background = t.btnHover; }}
            onMouseLeave={e => { e.currentTarget.style.color = t.textSecondary; e.currentTarget.style.background = isDark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.06)'; }}
          >
            {isDark ? Icon.sun : Icon.moon}
          </button>
        </div>
      </div>
    </aside>
  );
}

function HistoryRow({ item, active, t }: { item: HistoryItem; active: boolean; t: Theme }) {
  const dotColor = item.status === 'completed' ? '#10a37f' : item.status === 'failed' ? '#ef4444' : '#f59e0b';
  return (
    <button
      style={{ width: '100%', display: 'flex', alignItems: 'center', gap: 8, padding: '7px 10px', borderRadius: 7, background: active ? (t === DARK ? 'rgba(255,255,255,0.08)' : 'rgba(0,0,0,0.06)') : 'transparent', border: 'none', cursor: 'pointer', textAlign: 'left', transition: 'background 0.12s', marginBottom: 1 }}
      onMouseEnter={e => { if (!active) e.currentTarget.style.background = t.btnHover; }}
      onMouseLeave={e => { if (!active) e.currentTarget.style.background = 'transparent'; }}
    >
      <span style={{ width: 6, height: 6, borderRadius: '50%', background: dotColor, flexShrink: 0 }} />
      <span style={{ fontSize: 13, color: active ? t.text : t.textSecondary, flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
        {truncate(item.label)}
      </span>
    </button>
  );
}

// ─── Main ─────────────────────────────────────────────────────────────────────

export default function Home() {
  const [isDark, setIsDark] = useState(true);
  const t = isDark ? DARK : LIGHT;

  const [file, setFile] = useState<SelectedFile | null>(null);
  const [instruction, setInstruction] = useState('');
  const [stage, setStage] = useState<AppStage>('idle');
  const [jobId, setJobId] = useState<string | null>(null);
  const [jobResult, setJobResult] = useState<JobStatusResponse | null>(null);
  const [activityLog, setActivityLog] = useState<string[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [dragging, setDragging] = useState(false);
  const [history, setHistory] = useState<HistoryItem[]>([]);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const pollRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  // Load preferences from localStorage
  useEffect(() => {
    try {
      const saved = localStorage.getItem('excel-ai-history');
      if (saved) setHistory(JSON.parse(saved));
      const theme = localStorage.getItem('excel-ai-theme');
      if (theme) setIsDark(theme === 'dark');
    } catch {}
  }, []);

  const toggleTheme = () => {
    const next = !isDark;
    setIsDark(next);
    try { localStorage.setItem('excel-ai-theme', next ? 'dark' : 'light'); } catch {}
  };

  const saveHistory = (items: HistoryItem[]) => {
    setHistory(items);
    try { localStorage.setItem('excel-ai-history', JSON.stringify(items.slice(0, 30))); } catch {}
  };

  const startPolling = useCallback((id: string, instr: string) => {
    const poll = async () => {
      try {
        const res = await pollJobStatus(id);
        if (res.agentActivityLog?.length) setActivityLog(res.agentActivityLog);
        if (res.status === 'queued' || res.status === 'processing') {
          setStage(res.status === 'queued' ? 'queued' : 'processing');
          pollRef.current = setTimeout(poll, 2000);
        } else {
          setJobResult(res);
          const finalStatus = res.status === 'completed' ? 'completed'
            : res.status === 'clarification_needed' ? 'clarification_needed'
            : res.status === 'cannot_proceed' ? 'cannot_proceed' : 'failed';
          if (res.status === 'completed') setStage('completed');
          else if (res.status === 'clarification_needed') setStage('clarification');
          else { setStage('failed'); setError(res.error || res.cannotProceedReason || 'Processing failed'); }
          saveHistory([{ id, label: instr, timestamp: Date.now(), status: finalStatus as HistoryItem['status'] }, ...history]);
        }
      } catch (e: any) { setStage('failed'); setError(e.message); }
    };
    poll();
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [history]);

  useEffect(() => () => { if (pollRef.current) clearTimeout(pollRef.current); }, []);

  const acceptFile = useCallback((f: File) => {
    if (!f.name.match(/\.(xlsx|xlsm)$/i)) { setError('Only .xlsx files are supported.'); return; }
    if (f.size > 10 * 1024 * 1024) { setError('File must be under 10 MB.'); return; }
    setFile({ file: f, name: f.name, sizeBytes: f.size });
    setError(null);
    setStage(instruction.trim() ? 'ready' : 'idle');
    setTimeout(() => textareaRef.current?.focus(), 50);
  }, [instruction]);

  const onDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault(); setDragging(false);
    const f = e.dataTransfer.files[0]; if (f) acceptFile(f);
  }, [acceptFile]);

  const handleSubmit = async () => {
    if (!file || !instruction.trim() || isProcessing) return;
    setStage('submitting'); setError(null); setActivityLog([]); setJobResult(null);
    try {
      const { jobId: id } = await submitJob(file.file, instruction);
      setJobId(id); setStage('queued'); startPolling(id, instruction);
    } catch (e: any) { setStage('failed'); setError(e.message); }
  };

  const handleClarification = async (answer: string) => {
    if (!file || !answer.trim()) return;
    setStage('submitting'); setActivityLog([]); setJobResult(null);
    try {
      const { jobId: id } = await submitClarification(file.file, instruction, answer);
      setJobId(id); setStage('queued'); startPolling(id, instruction);
    } catch (e: any) { setStage('failed'); setError(e.message); }
  };

  const reset = () => {
    setFile(null); setInstruction(''); setStage('idle');
    setJobId(null); setJobResult(null); setActivityLog([]); setError(null);
  };

  const isProcessing = ['submitting', 'queued', 'processing'].includes(stage);
  const canSubmit = !!file && instruction.trim().length > 0 && !isProcessing;
  const isActive = stage !== 'idle';
  const statusLabel = stage === 'submitting' ? 'Submitting...' : stage === 'queued' ? 'In queue...' : 'Agent working...';

  return (
    <>
      <style>{`
        *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
        html, body { height: 100%; overflow: hidden; }
        ::placeholder { color: ${t.textMuted}; }
        :focus { outline: none; }
        textarea, input { font-family: inherit; color: ${t.text}; }
        ::-webkit-scrollbar { width: 5px; }
        ::-webkit-scrollbar-thumb { background: ${t.border}; border-radius: 3px; }
        @keyframes blink { 0%,100%{opacity:1} 50%{opacity:0} }
        @keyframes spin { to { transform: rotate(360deg); } }
        .spin { animation: spin 0.9s linear infinite; }
        .blink { animation: blink 1s step-end infinite; }
        body { background: ${t.bg}; color: ${t.text}; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; }
      `}</style>

      <div style={{ display: 'flex', height: '100vh', background: t.bg, overflow: 'hidden' }}>

        {/* ── Sidebar ── */}
        <Sidebar history={history} activeId={jobId} onNew={reset} isDark={isDark} onToggleTheme={toggleTheme} t={t} />

        {/* ── Main ── */}
        <div style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>

          {/* Top bar */}
          <div style={{ flexShrink: 0, padding: '14px 20px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', borderBottom: `1px solid ${t.borderSubtle}`, minHeight: 52 }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              {isActive && (
                <>
                  <div style={{ width: 6, height: 6, borderRadius: '50%', background: isProcessing ? '#f59e0b' : stage === 'completed' ? '#10a37f' : t.textMuted }} />
                  <span style={{ fontSize: 13, color: t.textSecondary }}>
                    {isProcessing ? statusLabel : stage === 'completed' ? 'Completed' : stage === 'clarification' ? 'Clarification needed' : stage === 'failed' ? 'Failed' : ''}
                  </span>
                </>
              )}
            </div>
            {isActive && (
              <button onClick={reset} style={{ fontSize: 13, color: t.textMuted, background: 'none', border: 'none', cursor: 'pointer', padding: '5px 10px', borderRadius: 6 }}
                onMouseEnter={e => (e.currentTarget.style.color = t.textSecondary)}
                onMouseLeave={e => (e.currentTarget.style.color = t.textMuted)}
              >New task</button>
            )}
          </div>

          {/* Content */}
          <div style={{ flex: 1, overflowY: 'auto', display: 'flex', flexDirection: 'column' }}>
            {!isActive ? (
              /* Idle */
              <div style={{ flex: 1, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', padding: '40px 24px 120px' }}>
                <h1 style={{ fontSize: 28, fontWeight: 500, color: t.text, marginBottom: 32, letterSpacing: '-0.02em' }}>
                  Ready when you are.
                </h1>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 10, width: '100%', maxWidth: 480 }}>
                  {['Update Management Fee to 1.5% on Fund Details sheet', 'Analyze the fee structure across all sheets', 'Add a new row for Q4 with value 25.3 in column C'].map((s, i) => (
                    <button key={i} onClick={() => { setInstruction(s); textareaRef.current?.focus(); }}
                      style={{ textAlign: 'left', padding: '10px 14px', borderRadius: 8, border: `1px solid ${t.border}`, background: t.card, color: t.textSecondary, fontSize: 13, cursor: 'pointer', transition: 'all 0.12s' }}
                      onMouseEnter={e => { e.currentTarget.style.borderColor = isDark ? 'rgba(255,255,255,0.18)' : 'rgba(0,0,0,0.2)'; e.currentTarget.style.color = t.text; }}
                      onMouseLeave={e => { e.currentTarget.style.borderColor = t.border; e.currentTarget.style.color = t.textSecondary; }}
                    >{s}</button>
                  ))}
                </div>
              </div>
            ) : (
              /* Active task */
              <div style={{ maxWidth: 680, margin: '0 auto', width: '100%', padding: '32px 24px', display: 'flex', flexDirection: 'column', gap: 20 }}>

                {/* User bubble */}
                <div style={{ alignSelf: 'flex-end', maxWidth: '80%', background: t.card, border: `1px solid ${t.border}`, borderRadius: '16px 16px 4px 16px', padding: '12px 16px' }}>
                  {file && (
                    <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8, padding: '6px 10px', background: t.bg, borderRadius: 7, border: `1px solid ${t.border}`, width: 'fit-content' }}>
                      <span style={{ color: '#10a37f' }}>{Icon.file}</span>
                      <span style={{ fontSize: 12, color: t.textSecondary }}>{file.name}</span>
                    </div>
                  )}
                  <p style={{ fontSize: 14, color: t.text, lineHeight: 1.5 }}>{instruction}</p>
                </div>

                {/* Agent response */}
                <div style={{ display: 'flex', gap: 12 }}>
                  <div style={{ width: 28, height: 28, borderRadius: '50%', background: t.card, border: `1px solid ${t.border}`, display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0, marginTop: 2 }}>
                    <svg width="14" height="14" fill="none" viewBox="0 0 24 24" stroke="#10a37f" strokeWidth={1.8}>
                      <path strokeLinecap="round" strokeLinejoin="round" d="M9.75 3.104v5.714a2.25 2.25 0 01-.659 1.591L5 14.5M9.75 3.104c-.251.023-.501.05-.75.082m.75-.082a24.301 24.301 0 014.5 0m0 0v5.714c0 .597.237 1.17.659 1.591L19.8 15.3M14.25 3.104c.251.023.501.05.75.082M19.8 15.3l-1.57.393A9.065 9.065 0 0112 15a9.065 9.065 0 00-6.23-.693L5 14.5m14.8.8l1.402 1.402c1 1 .03 2.798-1.442 2.798H4.24c-1.472 0-2.441-1.798-1.442-2.798L4.2 15.3"/>
                    </svg>
                  </div>
                  <div style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: 12 }}>

                    {isProcessing && (
                      <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        <div className="spin" style={{ width: 14, height: 14, border: `2px solid ${t.border}`, borderTopColor: '#10a37f', borderRadius: '50%' }} />
                        <span style={{ fontSize: 13, color: t.textSecondary }}>{statusLabel}</span>
                      </div>
                    )}

                    {(isProcessing || activityLog.length > 0) && (
                      <div>
                        <p style={{ fontSize: 11, color: t.textMuted, marginBottom: 6, textTransform: 'uppercase', letterSpacing: '0.06em', fontWeight: 500 }}>Agent log</p>
                        <ActivityLog lines={activityLog} running={isProcessing} t={t} />
                      </div>
                    )}

                    {error && (
                      <div style={{ display: 'flex', gap: 8, padding: '10px 14px', background: 'rgba(239,68,68,0.07)', border: '1px solid rgba(239,68,68,0.2)', borderRadius: 10 }}>
                        <span style={{ color: '#ef4444', flexShrink: 0, marginTop: 1 }}>{Icon.warn}</span>
                        <p style={{ fontSize: 13, color: '#fca5a5' }}>{error}</p>
                      </div>
                    )}

                    {stage === 'clarification' && jobResult && (
                      <ClarificationPanel question={jobResult.clarificationQuestion!} candidates={jobResult.clarificationCandidates ?? []} onAnswer={handleClarification} t={t} />
                    )}

                    {stage === 'completed' && jobResult && (
                      <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
                        <div style={{ display: 'flex', alignItems: 'flex-start', gap: 8 }}>
                          <span style={{ marginTop: 1 }}>{Icon.check}</span>
                          <div>
                            <p style={{ fontSize: 14, color: t.text, fontWeight: 500, marginBottom: 2 }}>Task completed</p>
                            <p style={{ fontSize: 13, color: t.textSecondary }}>{jobResult.operationSummary}</p>
                          </div>
                        </div>
                        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 8 }}>
                          {[
                            { label: 'Writes applied', value: jobResult.writeOpsApplied ?? '—' },
                            { label: 'Agent rounds', value: jobResult.agentRoundsUsed ?? '—' },
                            { label: 'Tool calls', value: jobResult.agentToolCallsMade ?? '—' },
                          ].map(({ label, value }) => (
                            <div key={label} style={{ background: t.bg, border: `1px solid ${t.borderSubtle}`, borderRadius: 8, padding: '10px 12px', textAlign: 'center' }}>
                              <p style={{ fontSize: 18, fontWeight: 600, color: t.text, marginBottom: 3 }}>{value}</p>
                              <p style={{ fontSize: 11, color: t.textMuted }}>{label}</p>
                            </div>
                          ))}
                        </div>
                        {jobResult.downloadUrl && (
                          <a href={`${process.env.NEXT_PUBLIC_API_URL || 'http://localhost:3001'}${jobResult.downloadUrl}`} download
                            style={{ display: 'inline-flex', alignItems: 'center', gap: 8, padding: '10px 18px', background: '#10a37f', color: '#fff', borderRadius: 8, fontSize: 13, fontWeight: 600, textDecoration: 'none', width: 'fit-content' }}
                            onMouseEnter={e => (e.currentTarget.style.background = '#0d9068')}
                            onMouseLeave={e => (e.currentTarget.style.background = '#10a37f')}
                          >
                            <svg width="14" height="14" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.5}><path strokeLinecap="round" strokeLinejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"/></svg>
                            Download file
                          </a>
                        )}
                      </div>
                    )}

                    {stage === 'failed' && jobResult?.status === 'cannot_proceed' && (
                      <div>
                        <p style={{ fontSize: 14, color: t.text, fontWeight: 500, marginBottom: 6 }}>Unable to proceed</p>
                        <p style={{ fontSize: 13, color: t.textSecondary, marginBottom: 6 }}>{jobResult.cannotProceedReason}</p>
                        <p style={{ fontSize: 12, color: t.textMuted }}>Try rephrasing your instruction.</p>
                      </div>
                    )}
                  </div>
                </div>
              </div>
            )}
          </div>

          {/* ── Bottom input bar ── */}
          <div style={{ flexShrink: 0, padding: '12px 20px 16px', borderTop: `1px solid ${t.borderSubtle}` }}>
            <div style={{ maxWidth: 680, margin: '0 auto' }}>

              {file && (
                <div style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '7px 14px', background: t.card, border: `1px solid ${t.border}`, borderRadius: 10, marginBottom: 8 }}>
                  <span style={{ color: '#10a37f' }}>{Icon.file}</span>
                  <span style={{ fontSize: 13, color: t.textSecondary, flex: 1 }}>{file.name}</span>
                  <span style={{ fontSize: 12, color: t.textMuted, marginRight: 8 }}>{formatBytes(file.sizeBytes)}</span>
                  {!isProcessing && (
                    <button onClick={() => { setFile(null); setStage('idle'); }} style={{ background: 'none', border: 'none', color: t.textMuted, cursor: 'pointer', display: 'flex', padding: 2, borderRadius: 4 }}
                      onMouseEnter={e => (e.currentTarget.style.color = t.textSecondary)}
                      onMouseLeave={e => (e.currentTarget.style.color = t.textMuted)}
                    >{Icon.x}</button>
                  )}
                </div>
              )}

              <div
                onDragOver={e => { e.preventDefault(); setDragging(true); }}
                onDragLeave={() => setDragging(false)}
                onDrop={onDrop}
                style={{ display: 'flex', alignItems: 'flex-end', gap: 8, padding: '10px 12px', background: t.input, border: `1px solid ${dragging ? '#10a37f' : t.border}`, borderRadius: 16, transition: 'border-color 0.15s' }}
              >
                <button onClick={() => fileInputRef.current?.click()} disabled={isProcessing}
                  style={{ width: 34, height: 34, borderRadius: 8, background: isDark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.05)', border: `1px solid ${t.border}`, display: 'flex', alignItems: 'center', justifyContent: 'center', cursor: isProcessing ? 'not-allowed' : 'pointer', flexShrink: 0, color: t.textSecondary, opacity: isProcessing ? 0.4 : 1 }}
                  title="Attach Excel file"
                >{Icon.plus}</button>

                <input ref={fileInputRef} type="file" accept=".xlsx,.xlsm" style={{ display: 'none' }} onChange={e => e.target.files?.[0] && acceptFile(e.target.files[0])} />

                <textarea
                  ref={textareaRef}
                  value={instruction}
                  onChange={e => { setInstruction(e.target.value); if (file && e.target.value.trim()) setStage('ready'); }}
                  onKeyDown={e => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); handleSubmit(); } }}
                  disabled={isProcessing}
                  placeholder={file ? 'Tell the agent what to do... (Enter to submit)' : 'Attach an Excel file first, then describe what to do...'}
                  rows={1}
                  style={{ flex: 1, background: 'transparent', border: 'none', fontSize: 14, color: t.text, resize: 'none', lineHeight: 1.5, paddingTop: 7, paddingBottom: 7, maxHeight: 120, opacity: isProcessing ? 0.5 : 1, overflowY: 'auto' }}
                  onInput={e => { const el = e.currentTarget; el.style.height = 'auto'; el.style.height = Math.min(el.scrollHeight, 120) + 'px'; }}
                />

                <button onClick={handleSubmit} disabled={!canSubmit}
                  style={{ width: 34, height: 34, borderRadius: '50%', background: canSubmit ? '#2563eb' : isDark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.05)', border: canSubmit ? 'none' : `1px solid ${t.border}`, display: 'flex', alignItems: 'center', justifyContent: 'center', cursor: canSubmit ? 'pointer' : 'not-allowed', flexShrink: 0, color: canSubmit ? '#fff' : t.textMuted, transition: 'background 0.15s' }}
                  title="Submit (Enter)"
                >
                  {isProcessing
                    ? <div className="spin" style={{ width: 14, height: 14, border: '2px solid rgba(255,255,255,0.2)', borderTopColor: '#fff', borderRadius: '50%' }} />
                    : Icon.send
                  }
                </button>
              </div>

              <p style={{ fontSize: 11, color: t.textMuted, textAlign: 'center', marginTop: 8 }}>
                Excel AI Agent — formula-safe, evidence-based automation
              </p>
            </div>
          </div>
        </div>
      </div>
    </>
  );
}
