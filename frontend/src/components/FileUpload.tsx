'use client';

import { useRef, useState, useCallback } from 'react';
import type { SelectedFile } from '@/types';

const ACCEPTED_MIME = [
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'application/vnd.ms-excel',
];
const ACCEPTED_EXT = ['.xlsx', '.xls'];
const MAX_SIZE_BYTES = 10 * 1024 * 1024; // 10 MB — matches backend

interface FileUploadProps {
  selected: SelectedFile | null;
  error?: string;
  disabled?: boolean;
  onSelect: (f: SelectedFile) => void;
  onClear: () => void;
  onError: (msg: string) => void;
}

export default function FileUpload({
  selected,
  error,
  disabled = false,
  onSelect,
  onClear,
  onError,
}: FileUploadProps) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [dragOver, setDragOver] = useState(false);

  const validate = useCallback(
    (file: File): string | null => {
      const ext = '.' + file.name.split('.').pop()?.toLowerCase();
      const validMime = ACCEPTED_MIME.includes(file.type);
      const validExt = ACCEPTED_EXT.includes(ext);

      if (!validMime && !validExt) {
        return 'Please select a valid Excel file (.xlsx or .xls).';
      }
      if (file.size > MAX_SIZE_BYTES) {
        return `File is too large. Maximum size is 10 MB (your file: ${formatBytes(file.size)}).`;
      }
      return null;
    },
    [],
  );

  const handleFile = useCallback(
    (file: File | null | undefined) => {
      if (!file) return;
      const err = validate(file);
      if (err) {
        onError(err);
        return;
      }
      onSelect({ file, name: file.name, sizeBytes: file.size });
    },
    [validate, onSelect, onError],
  );

  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    handleFile(e.target.files?.[0]);
    // Reset so the same file can be re-selected after clearing
    e.target.value = '';
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setDragOver(false);
    if (disabled) return;
    handleFile(e.dataTransfer.files?.[0]);
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    if (!disabled) setDragOver(true);
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if ((e.key === 'Enter' || e.key === ' ') && !disabled) {
      e.preventDefault();
      inputRef.current?.click();
    }
  };

  const hasError = !!error;

  return (
    <section aria-labelledby="upload-label">
      <label id="upload-label" style={styles.sectionLabel}>
        Excel File
      </label>

      {!selected ? (
        /* ── Empty / drop zone ─────────────────────────────────────────── */
        <div
          role="button"
          tabIndex={disabled ? -1 : 0}
          aria-disabled={disabled}
          aria-describedby={hasError ? 'upload-error' : 'upload-hint'}
          style={{
            ...styles.dropZone,
            ...(dragOver ? styles.dropZoneActive : {}),
            ...(hasError ? styles.dropZoneError : {}),
            ...(disabled ? styles.dropZoneDisabled : {}),
          }}
          onClick={() => !disabled && inputRef.current?.click()}
          onDrop={handleDrop}
          onDragOver={handleDragOver}
          onDragLeave={() => setDragOver(false)}
          onKeyDown={handleKeyDown}
        >
          <span style={styles.uploadIcon} aria-hidden="true">📂</span>
          <p style={styles.dropHeading}>
            {dragOver ? 'Drop file here' : 'Choose or drop an Excel file'}
          </p>
          <p id="upload-hint" style={styles.dropHint}>
            Accepts .xlsx and .xls · Max 10 MB
          </p>
          <button
            type="button"
            style={{
              ...styles.browseBtn,
              ...(disabled ? styles.browseBtnDisabled : {}),
            }}
            disabled={disabled}
            onClick={(e) => { e.stopPropagation(); inputRef.current?.click(); }}
            tabIndex={-1}
            aria-hidden="true"
          >
            Browse files
          </button>
        </div>
      ) : (
        /* ── File selected ─────────────────────────────────────────────── */
        <div style={styles.selectedCard} role="region" aria-label="Selected file">
          <div style={styles.fileInfo}>
            <span style={styles.fileIcon} aria-hidden="true">📊</span>
            <div style={styles.fileMeta}>
              <span style={styles.fileName}>{selected.name}</span>
              <span style={styles.fileSize}>{formatBytes(selected.sizeBytes)}</span>
            </div>
          </div>
          <button
            type="button"
            style={styles.changeBtn}
            onClick={() => { onClear(); inputRef.current?.click(); }}
            disabled={disabled}
            aria-label="Choose a different file"
          >
            Change file
          </button>
        </div>
      )}

      {/* Accessible error */}
      {hasError && (
        <p id="upload-error" role="alert" style={styles.errorText}>
          {error}
        </p>
      )}

      {/* Hidden file input */}
      <input
        ref={inputRef}
        type="file"
        accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
        onChange={handleInputChange}
        disabled={disabled}
        className="sr-only"
        aria-hidden="true"
        tabIndex={-1}
      />
    </section>
  );
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

// ─── Styles ───────────────────────────────────────────────────────────────────

const styles: Record<string, React.CSSProperties> = {
  sectionLabel: {
    display: 'block',
    fontWeight: 600,
    fontSize: '0.9rem',
    marginBottom: '0.6rem',
    color: 'var(--text)',
  },
  dropZone: {
    background: 'var(--surface)',
    border: '2px dashed var(--border)',
    borderRadius: 'var(--radius-lg)',
    padding: '2.5rem 1.5rem',
    textAlign: 'center',
    cursor: 'pointer',
    transition: 'border-color 0.15s, background 0.15s',
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    gap: '0.5rem',
  },
  dropZoneActive: {
    borderColor: 'var(--accent)',
    background: 'rgba(79,142,247,0.06)',
  },
  dropZoneError: {
    borderColor: 'var(--error)',
  },
  dropZoneDisabled: {
    opacity: 0.5,
    cursor: 'not-allowed',
  },
  uploadIcon: {
    fontSize: '2rem',
    lineHeight: 1,
  },
  dropHeading: {
    fontWeight: 500,
    fontSize: '0.95rem',
    color: 'var(--text)',
  },
  dropHint: {
    fontSize: '0.82rem',
    color: 'var(--text-muted)',
  },
  browseBtn: {
    marginTop: '0.75rem',
    background: 'transparent',
    border: '1px solid var(--border)',
    borderRadius: 'var(--radius)',
    color: 'var(--accent)',
    padding: '0.45rem 1.1rem',
    fontSize: '0.85rem',
    fontWeight: 500,
  },
  browseBtnDisabled: {
    opacity: 0.5,
    cursor: 'not-allowed',
  },
  selectedCard: {
    background: 'var(--surface)',
    border: '1px solid var(--border)',
    borderRadius: 'var(--radius-lg)',
    padding: '1rem 1.25rem',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: '1rem',
  },
  fileInfo: {
    display: 'flex',
    alignItems: 'center',
    gap: '0.75rem',
    minWidth: 0,
  },
  fileIcon: {
    fontSize: '1.5rem',
    flexShrink: 0,
  },
  fileMeta: {
    display: 'flex',
    flexDirection: 'column',
    minWidth: 0,
  },
  fileName: {
    fontWeight: 600,
    fontSize: '0.9rem',
    color: 'var(--text)',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  fileSize: {
    fontSize: '0.8rem',
    color: 'var(--text-muted)',
    marginTop: '0.1rem',
  },
  changeBtn: {
    background: 'transparent',
    border: '1px solid var(--border)',
    borderRadius: 'var(--radius)',
    color: 'var(--text-muted)',
    padding: '0.4rem 0.9rem',
    fontSize: '0.82rem',
    fontWeight: 500,
    flexShrink: 0,
    whiteSpace: 'nowrap',
  },
  errorText: {
    color: 'var(--error)',
    fontSize: '0.82rem',
    marginTop: '0.4rem',
  },
};
