/**
 * coordinate.utils.spec.ts
 *
 * Tests for the Excel A1 coordinate utility.
 *
 * All tests are deterministic and do not touch the filesystem.
 */

import {
  colLetterToNumber,
  colNumberToLetter,
  parseCell,
  buildCellAddress,
  parseRange,
  isCellAddressValid,
  isRangeValid,
  MAX_COL,
  MAX_ROW,
} from '../utils/coordinate.utils';

// ─── colLetterToNumber ────────────────────────────────────────────────────────

describe('colLetterToNumber', () => {
  describe('single letters', () => {
    it('A → 1', () => expect(colLetterToNumber('A')).toBe(1));
    it('B → 2', () => expect(colLetterToNumber('B')).toBe(2));
    it('Z → 26', () => expect(colLetterToNumber('Z')).toBe(26));
  });

  describe('double letters', () => {
    it('AA → 27', () => expect(colLetterToNumber('AA')).toBe(27));
    it('AB → 28', () => expect(colLetterToNumber('AB')).toBe(28));
    it('AZ → 52', () => expect(colLetterToNumber('AZ')).toBe(52));
    it('BA → 53', () => expect(colLetterToNumber('BA')).toBe(53));
  });

  describe('triple letters', () => {
    it('AAA → 703', () => expect(colLetterToNumber('AAA')).toBe(703));
    it('XFD → 16384 (MAX)', () => expect(colLetterToNumber('XFD')).toBe(16384));
  });

  describe('case insensitive', () => {
    it('lowercase a → 1', () => expect(colLetterToNumber('a')).toBe(1));
    it('mixed case Aa → 27', () => expect(colLetterToNumber('Aa')).toBe(27));
  });

  describe('invalid inputs', () => {
    it('empty string', () => expect(() => colLetterToNumber('')).toThrow());
    it('numbers only', () => expect(() => colLetterToNumber('1')).toThrow());
    it('letters+numbers', () => expect(() => colLetterToNumber('A1')).toThrow());
    it('spaces', () => expect(() => colLetterToNumber('A B')).toThrow());
    it('beyond MAX (XFE)', () => expect(() => colLetterToNumber('XFE')).toThrow());
  });
});

// ─── colNumberToLetter ────────────────────────────────────────────────────────

describe('colNumberToLetter', () => {
  it('1 → A', () => expect(colNumberToLetter(1)).toBe('A'));
  it('2 → B', () => expect(colNumberToLetter(2)).toBe('B'));
  it('26 → Z', () => expect(colNumberToLetter(26)).toBe('Z'));
  it('27 → AA', () => expect(colNumberToLetter(27)).toBe('AA'));
  it('28 → AB', () => expect(colNumberToLetter(28)).toBe('AB'));
  it('52 → AZ', () => expect(colNumberToLetter(52)).toBe('AZ'));
  it('53 → BA', () => expect(colNumberToLetter(53)).toBe('BA'));
  it('703 → AAA', () => expect(colNumberToLetter(703)).toBe('AAA'));
  it('16384 → XFD (MAX)', () => expect(colNumberToLetter(16384)).toBe('XFD'));

  describe('invalid inputs', () => {
    it('0', () => expect(() => colNumberToLetter(0)).toThrow());
    it('negative', () => expect(() => colNumberToLetter(-1)).toThrow());
    it('exceeds MAX', () => expect(() => colNumberToLetter(MAX_COL + 1)).toThrow());
    it('float', () => expect(() => colNumberToLetter(1.5)).toThrow());
  });
});

// ─── Round-trip conversions ───────────────────────────────────────────────────

describe('colLetter <-> colNumber round-trip', () => {
  const cases = [1, 7, 26, 27, 52, 53, 703, 16384];
  cases.forEach((n) => {
    it(`round-trip col ${n}`, () => {
      expect(colLetterToNumber(colNumberToLetter(n))).toBe(n);
    });
  });

  const letters = ['A', 'G', 'Z', 'AA', 'AZ', 'BA', 'AAA', 'XFD'];
  letters.forEach((l) => {
    it(`round-trip letter ${l}`, () => {
      expect(colNumberToLetter(colLetterToNumber(l))).toBe(l);
    });
  });
});

// ─── parseCell ────────────────────────────────────────────────────────────────

describe('parseCell', () => {
  it('A1 → row=1 col=1', () => {
    const r = parseCell('A1');
    expect(r.row).toBe(1);
    expect(r.col).toBe(1);
    expect(r.address).toBe('A1');
  });

  it('B1 → row=1 col=2', () => {
    const r = parseCell('B1');
    expect(r.row).toBe(1);
    expect(r.col).toBe(2);
    expect(r.address).toBe('B1');
  });

  it('Z1 → row=1 col=26', () => {
    const r = parseCell('Z1');
    expect(r.row).toBe(1);
    expect(r.col).toBe(26);
  });

  it('AA1 → row=1 col=27', () => {
    const r = parseCell('AA1');
    expect(r.row).toBe(1);
    expect(r.col).toBe(27);
  });

  it('AB1 → row=1 col=28', () => {
    const r = parseCell('AB1');
    expect(r.col).toBe(28);
  });

  it('AZ1 → row=1 col=52', () => {
    const r = parseCell('AZ1');
    expect(r.col).toBe(52);
  });

  it('BA1 → row=1 col=53', () => {
    const r = parseCell('BA1');
    expect(r.col).toBe(53);
  });

  it('G42 → row=42 col=7', () => {
    const r = parseCell('G42');
    expect(r.row).toBe(42);
    expect(r.col).toBe(7);
    expect(r.address).toBe('G42');
  });

  it('large valid: XFD1048576', () => {
    const r = parseCell('XFD1048576');
    expect(r.col).toBe(16384);
    expect(r.row).toBe(1048576);
  });

  describe('absolute references ($)', () => {
    it('$A$1 → row=1 col=1 address=A1', () => {
      const r = parseCell('$A$1');
      expect(r.row).toBe(1);
      expect(r.col).toBe(1);
      expect(r.address).toBe('A1'); // $ stripped in canonical address
    });

    it('$A1 → row=1 col=1', () => {
      const r = parseCell('$A1');
      expect(r.col).toBe(1);
    });

    it('A$1 → row=1 col=1', () => {
      const r = parseCell('A$1');
      expect(r.row).toBe(1);
    });
  });

  describe('case insensitive', () => {
    it('g42 = G42', () => {
      const r = parseCell('g42');
      expect(r.row).toBe(42);
      expect(r.col).toBe(7);
      expect(r.address).toBe('G42');
    });
  });

  describe('invalid inputs — must throw, not silently repair', () => {
    it('column only: A', () => expect(() => parseCell('A')).toThrow());
    it('row only: 42', () => expect(() => parseCell('42')).toThrow());
    it('digit before letter: 1A', () => expect(() => parseCell('1A')).toThrow());
    it('row zero: A0', () => expect(() => parseCell('A0')).toThrow());
    it('negative row: A-1', () => expect(() => parseCell('A-1')).toThrow());
    it('negative num: -1A', () => expect(() => parseCell('-1A')).toThrow());
    it('empty string', () => expect(() => parseCell('')).toThrow());
    it('spaces only', () => expect(() => parseCell('  ')).toThrow());
    it('beyond max row', () => expect(() => parseCell(`A${MAX_ROW + 1}`)).toThrow());
  });
});

// ─── buildCellAddress ─────────────────────────────────────────────────────────

describe('buildCellAddress', () => {
  it('(1,1) → A1', () => expect(buildCellAddress(1, 1)).toBe('A1'));
  it('(42,7) → G42', () => expect(buildCellAddress(42, 7)).toBe('G42'));
  it('(1,27) → AA1', () => expect(buildCellAddress(1, 27)).toBe('AA1'));

  it('invalid row 0', () => expect(() => buildCellAddress(0, 1)).toThrow());
  it('invalid col 0', () => expect(() => buildCellAddress(1, 0)).toThrow());
  it('negative row', () => expect(() => buildCellAddress(-1, 1)).toThrow());
});

// ─── parseRange ───────────────────────────────────────────────────────────────

describe('parseRange', () => {
  it('A1:C10 → rows=10 cols=3', () => {
    const r = parseRange('A1:C10');
    expect(r.rows).toBe(10);
    expect(r.cols).toBe(3);
    expect(r.start.address).toBe('A1');
    expect(r.end.address).toBe('C10');
    expect(r.address).toBe('A1:C10');
  });

  it('G42:J50 → rows=9 cols=4', () => {
    const r = parseRange('G42:J50');
    expect(r.rows).toBe(9);
    expect(r.cols).toBe(4);
  });

  it('AA5:AD20 → rows=16 cols=4', () => {
    const r = parseRange('AA5:AD20');
    expect(r.rows).toBe(16);
    expect(r.cols).toBe(4);
  });

  it('single cell A1:A1 → rows=1 cols=1', () => {
    const r = parseRange('A1:A1');
    expect(r.rows).toBe(1);
    expect(r.cols).toBe(1);
  });

  it('normalizes reversed range C10:A1 to A1:C10', () => {
    const r = parseRange('C10:A1');
    expect(r.start.address).toBe('A1');
    expect(r.end.address).toBe('C10');
    expect(r.rows).toBe(10);
    expect(r.cols).toBe(3);
  });

  it('absolute references: $A$1:$C$10', () => {
    const r = parseRange('$A$1:$C$10');
    expect(r.rows).toBe(10);
    expect(r.cols).toBe(3);
  });

  describe('invalid ranges — must throw', () => {
    it('missing colon: A1', () => expect(() => parseRange('A1')).toThrow());
    it('trailing colon: A1:', () => expect(() => parseRange('A1:')).toThrow());
    it('leading colon: :A10', () => expect(() => parseRange(':A10')).toThrow());
    it('row zero end: A1:B0', () => expect(() => parseRange('A1:B0')).toThrow());
    it('invalid start: 1A:B2', () => expect(() => parseRange('1A:B2')).toThrow());
    it('empty string', () => expect(() => parseRange('')).toThrow());
  });
});

// ─── Helper validators ────────────────────────────────────────────────────────

describe('isCellAddressValid', () => {
  it('A1 → true', () => expect(isCellAddressValid('A1')).toBe(true));
  it('G42 → true', () => expect(isCellAddressValid('G42')).toBe(true));
  it('A → false', () => expect(isCellAddressValid('A')).toBe(false));
  it('1A → false', () => expect(isCellAddressValid('1A')).toBe(false));
  it('A0 → false', () => expect(isCellAddressValid('A0')).toBe(false));
  it('empty → false', () => expect(isCellAddressValid('')).toBe(false));
});

describe('isRangeValid', () => {
  it('A1:C10 → true', () => expect(isRangeValid('A1:C10')).toBe(true));
  it('A1 → false (no colon)', () => expect(isRangeValid('A1')).toBe(false));
  it('A1: → false', () => expect(isRangeValid('A1:')).toBe(false));
  it(': → false', () => expect(isRangeValid(':')).toBe(false));
});
