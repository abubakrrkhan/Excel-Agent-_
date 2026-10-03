/**
 * financial-workbook.fixture.ts
 *
 * Generates a synthetic complex financial/fund-report Excel workbook.
 *
 * PURPOSE:
 * This fixture is designed to make naive coordinate assumptions FAIL.
 * It simulates realistic complexity found in bank/fund reports:
 *
 *   - Multiple worksheets with different structures
 *   - Merged headers (common in financial report layouts)
 *   - Repeated labels ("Management Fee" appears in multiple places)
 *   - Formulas and calculated values
 *   - Percentages, numbers, dates, text
 *   - Sparse cells
 *   - Footnotes
 *   - Cross-sheet references (as formula text — ExcelJS does not calculate)
 *
 * NOT based on any real bank's data.
 * Contains only synthetic values.
 *
 * This is the foundation for the future anti-hallucination test suite.
 */

import * as ExcelJS from 'exceljs';

// ─── Build the fixture ────────────────────────────────────────────────────────

/**
 * Build and return the fixture workbook.
 * Each call returns a fresh workbook instance.
 */
export function buildFinancialFixtureWorkbook(): ExcelJS.Workbook {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Excel AI Agent — Synthetic Fixture';
  wb.created = new Date('2024-01-01');

  addFundOverviewSheet(wb);
  addFundDetailsSheet(wb);
  addPortfolioSheet(wb);
  addFeeStructureSheet(wb);
  addFootnotesSheet(wb);

  return wb;
}

/**
 * Build and return the fixture as a Buffer (for use in tests).
 */
export async function buildFinancialFixtureBuffer(): Promise<Buffer> {
  const wb = buildFinancialFixtureWorkbook();
  const ab = await wb.xlsx.writeBuffer();
  return Buffer.from(ab);
}

// ─── Sheet 1: Fund Overview ───────────────────────────────────────────────────

function addFundOverviewSheet(wb: ExcelJS.Workbook): void {
  const ws = wb.addWorksheet('Fund Overview');

  // Row 1: Merged report title
  ws.mergeCells('A1:F1');
  ws.getCell('A1').value = 'Synthetic Capital Fund — Monthly Report';
  ws.getCell('A1').font = { bold: true, size: 14 };
  ws.getCell('A1').alignment = { horizontal: 'center' };

  // Row 2: Date row
  ws.getCell('A2').value = 'Report Date:';
  ws.getCell('B2').value = new Date('2024-01-31');
  ws.getCell('B2').numFmt = 'dd/mm/yyyy';

  ws.getCell('D2').value = 'Fund Manager:';
  ws.getCell('E2').value = 'Synthetic Asset Management';

  // Row 3: empty separator
  // Row 4: Key metrics header — merged
  ws.mergeCells('A4:F4');
  ws.getCell('A4').value = 'Key Fund Metrics';
  ws.getCell('A4').font = { bold: true };

  // Rows 5-10: Key metrics table
  const metrics = [
    ['Fund Size (PKR)', 2_500_000_000],
    ['NAV per Unit', 105.25],
    ['Management Fee', '1.25%'],        // NOTE: "Management Fee" appears here too
    ['Return (1M)', '2.35%'],
    ['Return (1Y)', '18.72%'],
    ['Benchmark Return (1Y)', '15.00%'],
  ];

  metrics.forEach(([label, value], i) => {
    const row = 5 + i;
    ws.getCell(`A${row}`).value = label as string;
    ws.getCell(`B${row}`).value = value;
  });

  // Row 12: empty
  // Row 13: Risk Metrics — deliberately reuses similar labels to test anti-hallucination
  ws.mergeCells('A13:F13');
  ws.getCell('A13').value = 'Risk & Cost Metrics';
  ws.getCell('A13').font = { bold: true };

  ws.getCell('A14').value = 'Management Fee';       // REPEATED LABEL — same text, different location
  ws.getCell('B14').value = '1.25%';
  ws.getCell('A15').value = 'Trustee Fee';
  ws.getCell('B15').value = '0.10%';
  ws.getCell('A16').value = 'Total Expense Ratio';
  ws.getCell('B16').value = '1.75%';

  // Footnote reference
  ws.getCell('A19').value = '* All percentages are annualized.';
  ws.getCell('A19').font = { italic: true, size: 9 };

  // Hide row 3 (spacer row).
  // ExcelJS only persists the hidden flag when the row has either content OR
  // explicit customHeight. We set both to guarantee serialization.
  const hiddenRow3 = ws.getRow(3);
  hiddenRow3.getCell(1).value = null; // empty cell forces row into the XML
  hiddenRow3.hidden = true;
  hiddenRow3.height = 0.1;

  // Hide column C (internal reference column not shown to users)
  ws.getColumn(3).hidden = true;
}

// ─── Sheet 2: Fund Details ────────────────────────────────────────────────────

function addFundDetailsSheet(wb: ExcelJS.Workbook): void {
  const ws = wb.addWorksheet('Fund Details');

  // Hide column E (internal calculation column — typical in financial reports)
  ws.getColumn(5).hidden = true;

  // Hide row 7 (audit/draft row — typical to hide work-in-progress rows)
  // This will be set after data is added below

  // Merged heading
  ws.mergeCells('A1:H1');
  ws.getCell('A1').value = 'Fund Structure & Fee Details';
  ws.getCell('A1').font = { bold: true, size: 13 };
  ws.getCell('A1').alignment = { horizontal: 'center' };

  // Column headers — row 3
  const headers = ['Parameter', 'Description', 'Value', 'Unit', '', 'Category', 'Benchmark', 'Notes'];
  headers.forEach((h, i) => {
    ws.getCell(3, i + 1).value = h;
    ws.getCell(3, i + 1).font = { bold: true };
  });

  // Merged header for categories
  ws.mergeCells('F3:G3');
  ws.getCell('F3').value = 'Performance';
  ws.getCell('F3').alignment = { horizontal: 'center' };

  // Data rows — rows 5–20
  const rows: Array<[string, string, string | number, string]> = [
    ['Fund Name', 'Full legal name', 'Synthetic Capital Fund', 'N/A'],
    ['Fund Type', 'Regulatory category', 'Open End', 'N/A'],
    ['Launch Date', 'Date of inception', '01/03/2018', 'Date'],
    ['Fiscal Year', 'Fund fiscal year end', 'June 30', 'N/A'],
    ['', '', '', ''],                               // sparse row
    ['Management Fee', 'Annual fee on NAV', 1.25, '%'],    // THE TARGET cell — row 10 in this section
    ['Performance Fee', 'Fee on excess returns', 10.00, '% of excess'],
    ['Trustee Fee', 'Custodian/trustee', 0.10, '%'],
    ['Sales Load (Front)', 'Entry load', 0.00, '%'],
    ['Sales Load (Back)', 'Exit load', 0.00, '%'],
    ['', '', '', ''],                               // sparse row
    ['Min Investment (PKR)', 'Minimum initial', 10_000, 'PKR'],
    ['Min Subsequent (PKR)', 'Minimum top-up', 5_000, 'PKR'],
    ['Redemption Period', 'T+2 settlement', 2, 'Business Days'],
    ['', '', '', ''],                               // sparse row
    ['Total AUM (PKR M)', 'Assets under mgmt', 2_500, 'PKR M'],
  ];

  rows.forEach((rowData, i) => {
    const rowNum = 5 + i;
    rowData.forEach((val, ci) => {
      ws.getCell(rowNum, ci + 1).value = val as any;
    });
  });

  // The critical cell for future anti-hallucination tests:
  // Management Fee = G10 (row 10, col 3 = C, but let's use actual row numbering)
  // Row 5+5=10, col 3 = C — value is 1.25 (numeric), col 4 = "%" unit
  // So "Management Fee" text is in C10, value 1.25 is in C10... wait let me be precise

  // Row 5 = Fund Name, Row 6 = Fund Type, ..., Row 10 = Management Fee (5+5=10)
  // Col A=1 = "Management Fee" label, Col C=3 = 1.25 value
  // This makes G42 (from the context docs) NOT the correct cell — as intended.
  // The fixture is designed so no simple assumption works.

  // Add a formula in a later row
  ws.getCell('C21').value = { formula: 'SUM(C10:C12)', result: 11.35 };
  ws.getCell('A21').value = 'Total Fee Burden';
  ws.getCell('C21').numFmt = '0.00"%"';

  // Footnote
  ws.getCell('A23').value = '** Management Fee subject to annual review.';
  ws.getCell('A23').font = { italic: true, size: 9 };

  // ── Second section (rows 26–32): Performance Data ─────────────────────────
  // Separated from the fee section by blank rows 24–25
  // This tests multi-section detection
  ws.mergeCells('A26:D26');
  ws.getCell('A26').value = 'Historical Performance';
  ws.getCell('A26').font = { bold: true };

  const perfHeaders = ['Period', 'Fund Return', 'Benchmark', 'Excess Return'];
  perfHeaders.forEach((h, i) => {
    ws.getCell(27, i + 1).value = h;
    ws.getCell(27, i + 1).font = { bold: true };
  });

  const perfRows = [
    ['1 Month', '2.35%', '1.90%', '0.45%'],
    ['3 Months', '6.80%', '5.60%', '1.20%'],
    ['6 Months', '11.20%', '9.80%', '1.40%'],
    ['1 Year', '18.72%', '15.00%', '3.72%'],
  ];
  perfRows.forEach((row, i) => {
    row.forEach((val, ci) => ws.getCell(28 + i, ci + 1).value = val);
  });

  // Formula for average excess return
  ws.getCell('D32').value = { formula: 'AVERAGE(D28:D31)', result: 1.6925 };
  ws.getCell('A32').value = 'Average Excess';
  ws.getCell('A32').font = { bold: true };

  // Now hide row 7 (internal audit row in the fee section)
  ws.getRow(7).hidden = true;
}

// ─── Sheet 3: Portfolio ───────────────────────────────────────────────────────

function addPortfolioSheet(wb: ExcelJS.Workbook): void {
  const ws = wb.addWorksheet('Portfolio');

  // Merged sector header
  ws.mergeCells('A1:E1');
  ws.getCell('A1').value = 'Portfolio Holdings — January 2024';
  ws.getCell('A1').font = { bold: true, size: 12 };

  // Table headers
  ['Security', 'Sector', 'Weight %', 'Quantity', 'Market Value (PKR)'].forEach(
    (h, i) => {
      const cell = ws.getCell(2, i + 1);
      cell.value = h;
      cell.font = { bold: true };
    },
  );

  // Merged sector sub-headers (simulates financial report grouping)
  ws.mergeCells('A3:E3');
  ws.getCell('A3').value = 'Equity Holdings';
  ws.getCell('A3').font = { italic: true, bold: true };

  const equities = [
    ['HBL', 'Banking', 8.5, 50000, 12_750_000],
    ['MCB', 'Banking', 6.2, 40000, 9_300_000],
    ['ENGRO', 'Fertilizer', 5.1, 30000, 7_650_000],
    ['LUCK', 'Cement', 4.8, 20000, 7_200_000],
    ['OGDC', 'Energy', 7.3, 60000, 10_950_000],
  ];

  equities.forEach((row, i) => {
    row.forEach((val, ci) => {
      ws.getCell(4 + i, ci + 1).value = val as any;
    });
    // Format weight as percentage
    ws.getCell(4 + i, 3).numFmt = '0.00%';
  });

  // Sub-total row with formula
  ws.getCell('A9').value = 'Equity Sub-Total';
  ws.getCell('A9').font = { bold: true };
  ws.getCell('C9').value = { formula: 'SUM(C4:C8)', result: 31.9 };

  // Rows 10–11: explicit blank separator rows (2-row gap triggers section boundary)
  ws.getRow(10).height = 8;
  ws.getRow(11).height = 8;

  // Fixed income section starts at row 12
  ws.mergeCells('A12:E12');
  ws.getCell('A12').value = 'Fixed Income Holdings';
  ws.getCell('A12').font = { italic: true, bold: true };

  const fixedIncome = [
    ['PIB 5Y', 'Government', 15.0, 100, 37_500_000],
    ['T-Bill 3M', 'Government', 10.5, 200, 26_250_000],
    ['TFC - ABC Bank', 'Corporate Bond', 8.2, 50, 20_500_000],
  ];

  fixedIncome.forEach((row, i) => {
    row.forEach((val, ci) => {
      ws.getCell(13 + i, ci + 1).value = val as any;
    });
  });

  ws.getCell('A16').value = 'Fixed Income Sub-Total';
  ws.getCell('A16').font = { bold: true };
  ws.getCell('C16').value = { formula: 'SUM(C13:C15)', result: 33.7 };

  // Cash
  ws.getCell('A18').value = 'Cash & Equivalents';
  ws.getCell('C18').value = 34.4;

  // Grand total
  ws.getCell('A20').value = 'Total Portfolio';
  ws.getCell('A20').font = { bold: true };
  ws.getCell('C20').value = { formula: 'C9+C16+C18', result: 100.0 };
  ws.getCell('C20').numFmt = '0.00"%"';
}

// ─── Sheet 4: Fee Structure ───────────────────────────────────────────────────

function addFeeStructureSheet(wb: ExcelJS.Workbook): void {
  const ws = wb.addWorksheet('Fee Structure');

  // ANOTHER sheet with "Management Fee" — tests multi-location search
  ws.mergeCells('A1:D1');
  ws.getCell('A1').value = 'Fee Structure — All Share Classes';
  ws.getCell('A1').font = { bold: true, size: 12 };

  // Headers
  ['Fee Type', 'Class A', 'Class B', 'Class C'].forEach((h, i) => {
    ws.getCell(3, i + 1).value = h;
    ws.getCell(3, i + 1).font = { bold: true };
  });

  const fees = [
    ['Management Fee', '1.25%', '1.00%', '0.75%'],  // "Management Fee" — 3rd occurrence across workbook
    ['Performance Fee', '10%', '10%', '5%'],
    ['Trustee Fee', '0.10%', '0.10%', '0.10%'],
    ['Front-end Load', '0.00%', '0.00%', '0.00%'],
    ['Back-end Load', '0.00%', '0.00%', '0.00%'],
    ['Minimum Investment', 'PKR 10,000', 'PKR 500,000', 'PKR 5,000,000'],
  ];

  fees.forEach((row, i) => {
    row.forEach((val, ci) => {
      ws.getCell(4 + i, ci + 1).value = val;
    });
  });

  ws.getCell('A11').value = 'Note: Class B and C require minimum 1-year lock-in.';
  ws.getCell('A11').font = { italic: true, size: 9 };
}

// ─── Sheet 5: Footnotes ───────────────────────────────────────────────────────

function addFootnotesSheet(wb: ExcelJS.Workbook): void {
  const ws = wb.addWorksheet('Notes');

  ws.getCell('A1').value = 'Disclosures and Footnotes';
  ws.getCell('A1').font = { bold: true, size: 12 };

  const notes = [
    '1. Past performance is not indicative of future results.',
    '2. NAV is calculated daily on the basis of net assets.',
    '3. Management fee as disclosed in the offering document.',
    '4. All return calculations are on NAV-to-NAV basis.',
    '5. Total Expense Ratio (TER) = Management Fee + Trustee Fee + Other Expenses.',
    '6. This is a synthetic fixture workbook. No real financial data is contained.',
  ];

  notes.forEach((note, i) => {
    ws.getCell(`A${3 + i}`).value = note;
  });

  // Sparse area — a few cells scattered
  ws.getCell('D15').value = 'Cross-ref: Fund Details!C10';  // pseudo cross-sheet reference
  ws.getCell('D16').value = 'Management Fee rate effective from 2022-01-01.'; // 4th occurrence
}
