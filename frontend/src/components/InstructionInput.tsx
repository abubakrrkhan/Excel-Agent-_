'use client';

const MAX_CHARS = 2000;

const EXAMPLES = [
  'Set Alice\'s salary to 100000',
  'Add a "Total" column with a SUM formula',
  'Sort the sheet by Salary in descending order',
  'Delete all rows where Department is "Temp"',
  'Format the header row in bold',
];

interface InstructionInputProps {
  value: string;
  error?: string;
  disabled?: boolean;
  onChange: (value: string) => void;
}

export default function InstructionInput({
  value,
  error,
  disabled = false,
  onChange,
}: InstructionInputProps) {
  const remaining = MAX_CHARS - value.length;
  const hasError = !!error;
  const isNearLimit = remaining < 200;
  const isAtLimit = remaining <= 0;

  return (
    <section aria-labelledby="instruction-label">
      <div style={styles.labelRow}>
        <label id="instruction-label" htmlFor="instruction-input" style={styles.label}>
          What would you like to change?
        </label>
        <span
          style={{
            ...styles.charCount,
            color: isAtLimit
              ? 'var(--error)'
              : isNearLimit
              ? 'var(--warning)'
              : 'var(--text-subtle)',
          }}
          aria-live="polite"
          aria-label={`${remaining} characters remaining`}
        >
          {remaining.toLocaleString()}
        </span>
      </div>

      <textarea
        id="instruction-input"
        value={value}
        onChange={(e) => onChange(e.target.value.slice(0, MAX_CHARS))}
        disabled={disabled}
        rows={4}
        placeholder={EXAMPLES[0]}
        maxLength={MAX_CHARS}
        aria-describedby={hasError ? 'instruction-error' : 'instruction-hint'}
        aria-invalid={hasError}
        style={{
          ...styles.textarea,
          ...(hasError ? styles.textareaError : {}),
          ...(disabled ? styles.textareaDisabled : {}),
        }}
      />

      {hasError ? (
        <p id="instruction-error" role="alert" style={styles.error}>
          {error}
        </p>
      ) : (
        <p id="instruction-hint" style={styles.hint}>
          Describe the change in plain English. Examples:{' '}
          <em style={{ color: 'var(--text-muted)' }}>
            &ldquo;{EXAMPLES[1]}&rdquo;
          </em>
        </p>
      )}
    </section>
  );
}

const styles: Record<string, React.CSSProperties> = {
  labelRow: {
    display: 'flex',
    alignItems: 'baseline',
    justifyContent: 'space-between',
    marginBottom: '0.6rem',
  },
  label: {
    fontWeight: 600,
    fontSize: '0.9rem',
    color: 'var(--text)',
  },
  charCount: {
    fontSize: '0.78rem',
    fontVariantNumeric: 'tabular-nums',
    transition: 'color 0.2s',
  },
  textarea: {
    width: '100%',
    background: 'var(--surface)',
    border: '1px solid var(--border)',
    borderRadius: 'var(--radius)',
    color: 'var(--text)',
    padding: '0.75rem 1rem',
    fontSize: '0.95rem',
    lineHeight: 1.6,
    resize: 'vertical',
    fontFamily: 'inherit',
    display: 'block',
    minHeight: 100,
    transition: 'border-color 0.15s, box-shadow 0.15s',
  },
  textareaError: {
    borderColor: 'var(--error)',
  },
  textareaDisabled: {
    opacity: 0.5,
    cursor: 'not-allowed',
    resize: 'none',
  },
  hint: {
    fontSize: '0.8rem',
    color: 'var(--text-subtle)',
    marginTop: '0.4rem',
  },
  error: {
    color: 'var(--error)',
    fontSize: '0.82rem',
    marginTop: '0.4rem',
  },
};
