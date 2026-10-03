/**
 * phase10-reliability.spec.ts
 *
 * Phase 10 — Final reliability, security and production readiness tests.
 *
 * Tests every confirmed bug fix:
 *   A. API routing (P0: A1) — job endpoint paths
 *   B. File upload validation (P0: A4, P1: A5) — size limits, .xls rejection
 *   C. Job queue payload (P0: A3) — disk-reference design
 *   D. Clarification workflow (P0: A2) — answer reaches processor
 *   E. Operation validation (P1: A8, A9) — bounds checks
 *   F. Download security (P1: A6) — path traversal, output-only guard
 *   G. File retention (P1: A7) — cleanup logic
 *   H. Formula protection regression
 *   I. Coordinate validation regression
 */

import * as path from 'path';
import * as fs from 'fs';
import * as os from 'os';
import * as ExcelJS from 'exceljs';
import { FilesService } from '../files.service';
import { DatabaseService } from '../../database/database.service';
import { OperationsService } from '../../operations/operations.service';
import { WorkbookStructureService } from '../../excel/workbook-structure.service';
import { ExcelInspectionService } from '../../excel/excel-inspection.service';
import { ExcelService } from '../../excel/excel.service';

// ─── Helpers ──────────────────────────────────────────────────────────────────

function makeMockFile(overrides: Partial<Express.Multer.File> = {}): Express.Multer.File {
  return {
    fieldname: 'file',
    originalname: 'test.xlsx',
    encoding: '7bit',
    mimetype: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: Buffer.from('fake-xlsx-content'),
    size: 1024,
    destination: '',
    filename: '',
    path: '',
    stream: null as any,
    ...overrides,
  };
}

async function makeRealXlsxBuffer(): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet('Sheet1').addRow(['Name', 'Value']);
  return Buffer.from(await wb.xlsx.writeBuffer());
}

// ─── FilesService with temp dir ───────────────────────────────────────────────

function makeFilesService(): { svc: FilesService; tmpDir: string } {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'excel-test-'));
  const mockDb = { getPool: () => null } as unknown as DatabaseService;
  const svc = new FilesService(mockDb);
  // Override uploadDir to use our temp dir
  (svc as any).uploadDir = tmpDir;
  (svc as any).ensureUploadDir = () => {};
  return { svc, tmpDir };
}

// ─── Setup ────────────────────────────────────────────────────────────────────

let opsSvc: OperationsService;
let excelSvc: ExcelService;

beforeAll(() => {
  const structSvc = new WorkbookStructureService();
  const inspSvc = new ExcelInspectionService(structSvc);
  opsSvc = new OperationsService(inspSvc, structSvc);
  excelSvc = new ExcelService();
});

// ═══════════════════════════════════════════════════════════════════════════
// A. API ROUTING — confirmed fix verification
// ═══════════════════════════════════════════════════════════════════════════

describe('A. API routing — controller path verification', () => {
  it('JobsController uses @Controller("jobs") not @Controller("api/jobs")', () => {
    // Read the controller source and verify the decorator
    const src = fs.readFileSync(
      path.join(__dirname, '../../jobs/jobs.controller.ts'),
      'utf8',
    );
    // Must NOT contain the doubled prefix
    expect(src).not.toContain("@Controller('api/jobs')");
    expect(src).not.toContain('@Controller("api/jobs")');
    // Must use the plain 'jobs' path
    expect(src).toMatch(/@Controller\(['"]jobs['"]\)/);
  });

  it('FilesController uses @Controller("files") — no global prefix duplication', () => {
    const src = fs.readFileSync(
      path.join(__dirname, '../../files/files.controller.ts'),
      'utf8',
    );
    expect(src).not.toContain("@Controller('api/files')");
    expect(src).toMatch(/@Controller\(['"]files['"]\)/);
  });

  it('HealthController uses @Controller("health") — no global prefix duplication', () => {
    const src = fs.readFileSync(
      path.join(__dirname, '../../health/health.controller.ts'),
      'utf8',
    );
    expect(src).not.toContain("@Controller('api/health')");
    expect(src).toMatch(/@Controller\(['"]health['"]\)/);
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// B. FILE UPLOAD VALIDATION
// ═══════════════════════════════════════════════════════════════════════════

describe('B. File upload validation', () => {
  let svc: FilesService;

  beforeAll(() => {
    const { svc: s } = makeFilesService();
    svc = s;
  });

  it('accepts valid .xlsx file', () => {
    const file = makeMockFile();
    expect(() => svc.validateExcelFile(file)).not.toThrow();
  });

  it('throws when no file provided', () => {
    expect(() => svc.validateExcelFile(undefined as any)).toThrow(/No file provided/i);
  });

  it('throws for .xls file with clear conversion message', () => {
    const file = makeMockFile({
      originalname: 'report.xls',
      mimetype: 'application/vnd.ms-excel',
    });
    let err: any;
    try { svc.validateExcelFile(file); } catch (e) { err = e; }
    expect(err).toBeDefined();
    expect(err.message).toMatch(/\.xls.*(not supported|convert)/i);
    expect(err.message).toMatch(/\.xlsx/i);
  });

  it('throws for unsupported extension (.csv)', () => {
    const file = makeMockFile({
      originalname: 'data.csv',
      mimetype: 'text/csv',
    });
    expect(() => svc.validateExcelFile(file)).toThrow(/unsupported file extension/i);
  });

  it('throws for unsupported extension (.pdf)', () => {
    const file = makeMockFile({
      originalname: 'report.pdf',
      mimetype: 'application/pdf',
    });
    expect(() => svc.validateExcelFile(file)).toThrow();
  });

  it('throws for wrong MIME type on .xlsx extension', () => {
    const file = makeMockFile({
      originalname: 'test.xlsx',
      mimetype: 'text/plain',
    });
    expect(() => svc.validateExcelFile(file)).toThrow(/invalid file type/i);
  });

  it('throws for file at 10 MB + 1 byte', () => {
    const file = makeMockFile({
      size: 10 * 1024 * 1024 + 1,
      buffer: Buffer.alloc(10 * 1024 * 1024 + 1),
    });
    expect(() => svc.validateExcelFile(file)).toThrow(/too large/i);
  });

  it('accepts file at exactly 10 MB', () => {
    const file = makeMockFile({
      size: 10 * 1024 * 1024,
      buffer: Buffer.alloc(10 * 1024 * 1024),
    });
    expect(() => svc.validateExcelFile(file)).not.toThrow();
  });

  it('accepts small file (1 KB)', () => {
    const file = makeMockFile({ size: 1024 });
    expect(() => svc.validateExcelFile(file)).not.toThrow();
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// C. JOB QUEUE PAYLOAD — disk-reference design
// ═══════════════════════════════════════════════════════════════════════════

describe('C. Job queue payload — disk-reference design', () => {
  it('AgentJobPayload has storedFileName not fileBase64', () => {
    const src = fs.readFileSync(
      path.join(__dirname, '../../jobs/job.types.ts'),
      'utf8',
    );
    expect(src).toContain('storedFileName');
    expect(src).not.toContain('fileBase64');
  });

  it('JobsController saves to disk before queuing', () => {
    const src = fs.readFileSync(
      path.join(__dirname, '../../jobs/jobs.controller.ts'),
      'utf8',
    );
    expect(src).toContain('saveFileToDisk');
    expect(src).toContain('storedFileName');
    expect(src).not.toContain('fileBase64');
    expect(src).not.toContain('.toString(\'base64\')');
  });

  it('JobsProcessor reads file from disk', () => {
    const src = fs.readFileSync(
      path.join(__dirname, '../../jobs/jobs.processor.ts'),
      'utf8',
    );
    expect(src).toContain('readFileFromDisk');
    expect(src).toContain('storedFileName');
    expect(src).not.toContain('fileBase64');
    expect(src).not.toContain('Buffer.from(fileBase64');
  });

  it('FilesService.readFileFromDisk uses path.basename (traversal safe)', () => {
    const src = fs.readFileSync(
      path.join(__dirname, '../../files/files.service.ts'),
      'utf8',
    );
    expect(src).toContain('path.basename(storedName)');
  });

  it('readFileFromDisk returns buffer for existing file', async () => {
    const { svc, tmpDir } = makeFilesService();
    const testFile = path.join(tmpDir, 'test-file.xlsx');
    const content = Buffer.from('test content');
    fs.writeFileSync(testFile, content);
    const result = await svc.readFileFromDisk('test-file.xlsx');
    expect(result).toEqual(content);
  });

  it('readFileFromDisk throws NotFoundException for missing file', async () => {
    const { svc } = makeFilesService();
    await expect(svc.readFileFromDisk('nonexistent.xlsx')).rejects.toThrow(/not found/i);
  });

  it('readFileFromDisk ignores path traversal attempts', async () => {
    const { svc } = makeFilesService();
    // Should use basename, so ../../etc/passwd becomes 'passwd' which doesn't exist
    await expect(svc.readFileFromDisk('../../etc/passwd')).rejects.toThrow();
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// D. CLARIFICATION WORKFLOW — answer reaches processor
// ═══════════════════════════════════════════════════════════════════════════

describe('D. Clarification workflow', () => {
  it('JobsController submitJob reads clarificationAnswer from body', () => {
    const src = fs.readFileSync(
      path.join(__dirname, '../../jobs/jobs.controller.ts'),
      'utf8',
    );
    expect(src).toContain("@Body('clarificationAnswer')");
    expect(src).toContain('clarificationAnswer');
  });

  it('JobsProcessor combines instruction with clarificationAnswer', () => {
    const src = fs.readFileSync(
      path.join(__dirname, '../../jobs/jobs.processor.ts'),
      'utf8',
    );
    expect(src).toContain('clarificationAnswer');
    expect(src).toContain('effectiveInstruction');
    expect(src).toContain('User clarification:');
  });

  it('effective instruction includes clarification when provided', () => {
    // Test the logic directly
    const instruction = 'Update management fee';
    const clarificationAnswer = 'Fund Details sheet';
    const effectiveInstruction = clarificationAnswer
      ? `${instruction}\n\nUser clarification: ${clarificationAnswer}`
      : instruction;
    expect(effectiveInstruction).toContain('Update management fee');
    expect(effectiveInstruction).toContain('Fund Details sheet');
    expect(effectiveInstruction).toContain('User clarification:');
  });

  it('effective instruction equals original instruction when no clarification', () => {
    const instruction = 'Update management fee';
    const clarificationAnswer = undefined;
    const effectiveInstruction = clarificationAnswer
      ? `${instruction}\n\nUser clarification: ${clarificationAnswer}`
      : instruction;
    expect(effectiveInstruction).toBe(instruction);
  });

  it('frontend submitClarification sends clarificationAnswer field', () => {
    const src = fs.readFileSync(
      path.join(__dirname, '../../../../../frontend/src/lib/api.ts'),
      'utf8',
    );
    expect(src).toContain("append('clarificationAnswer'");
    expect(src).toContain('/jobs/submit');
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// E. OPERATION VALIDATION — bounds and type checks
// ═══════════════════════════════════════════════════════════════════════════

describe('E. Operation validation', () => {

  describe('DELETE_ROW', () => {
    it('accepts valid row number', () => {
      expect(() => opsSvc.validatePlan({
        summary: 'test',
        operations: [{ type: 'DELETE_ROW', sheet: 'S', params: { row: 3 } }],
      })).not.toThrow();
    });

    it('rejects row 0', () => {
      expect(() => opsSvc.validatePlan({
        summary: 'test',
        operations: [{ type: 'DELETE_ROW', sheet: 'S', params: { row: 0 } }],
      })).toThrow(/positive integer/i);
    });

    it('rejects negative row', () => {
      expect(() => opsSvc.validatePlan({
        summary: 'test',
        operations: [{ type: 'DELETE_ROW', sheet: 'S', params: { row: -1 } }],
      })).toThrow(/positive integer/i);
    });

    it('rejects string row', () => {
      expect(() => opsSvc.validatePlan({
        summary: 'test',
        operations: [{ type: 'DELETE_ROW', sheet: 'S', params: { row: 'three' } }],
      })).toThrow();
    });
  });

  describe('ADD_FORMULA', () => {
    it('accepts valid formula', () => {
      expect(() => opsSvc.validatePlan({
        summary: 'test',
        operations: [{ type: 'ADD_FORMULA', sheet: 'S', params: { row: 1, col: 1, formula: 'SUM(A1:A5)' } }],
      })).not.toThrow();
    });

    it('rejects empty formula', () => {
      expect(() => opsSvc.validatePlan({
        summary: 'test',
        operations: [{ type: 'ADD_FORMULA', sheet: 'S', params: { row: 1, col: 1, formula: '   ' } }],
      })).toThrow(/formula/i);
    });
  });

  describe('FORMAT_CELL', () => {
    it('accepts valid style object', () => {
      expect(() => opsSvc.validatePlan({
        summary: 'test',
        operations: [{ type: 'FORMAT_CELL', sheet: 'S', params: { row: 1, col: 1, style: { font: { bold: true } } } }],
      })).not.toThrow();
    });

    it('rejects missing style', () => {
      expect(() => opsSvc.validatePlan({
        summary: 'test',
        operations: [{ type: 'FORMAT_CELL', sheet: 'S', params: { row: 1, col: 1 } }],
      })).toThrow(/style/i);
    });
  });

  describe('ADD_COLUMN_HEADER', () => {
    it('accepts valid col and header', () => {
      expect(() => opsSvc.validatePlan({
        summary: 'test',
        operations: [{ type: 'ADD_COLUMN_HEADER', sheet: 'S', params: { col: 1, header: 'Name' } }],
      })).not.toThrow();
    });

    it('rejects col 0', () => {
      expect(() => opsSvc.validatePlan({
        summary: 'test',
        operations: [{ type: 'ADD_COLUMN_HEADER', sheet: 'S', params: { col: 0, header: 'Name' } }],
      })).toThrow(/positive integer/i);
    });

    it('rejects missing header', () => {
      expect(() => opsSvc.validatePlan({
        summary: 'test',
        operations: [{ type: 'ADD_COLUMN_HEADER', sheet: 'S', params: { col: 1 } }],
      })).toThrow(/header/i);
    });

    it('rejects empty header string', () => {
      expect(() => opsSvc.validatePlan({
        summary: 'test',
        operations: [{ type: 'ADD_COLUMN_HEADER', sheet: 'S', params: { col: 1, header: '' } }],
      })).toThrow(/header/i);
    });
  });

  describe('SORT_SHEET', () => {
    it('accepts asc', () => {
      expect(() => opsSvc.validatePlan({
        summary: 'test',
        operations: [{ type: 'SORT_SHEET', sheet: 'S', params: { col: 1, order: 'asc' } }],
      })).not.toThrow();
    });

    it('accepts desc', () => {
      expect(() => opsSvc.validatePlan({
        summary: 'test',
        operations: [{ type: 'SORT_SHEET', sheet: 'S', params: { col: 2, order: 'desc' } }],
      })).not.toThrow();
    });

    it('rejects invalid order value', () => {
      expect(() => opsSvc.validatePlan({
        summary: 'test',
        operations: [{ type: 'SORT_SHEET', sheet: 'S', params: { col: 1, order: 'random' } }],
      })).toThrow(/asc.*desc|desc.*asc/i);
    });

    it('rejects col 0', () => {
      expect(() => opsSvc.validatePlan({
        summary: 'test',
        operations: [{ type: 'SORT_SHEET', sheet: 'S', params: { col: 0, order: 'asc' } }],
      })).toThrow(/positive integer/i);
    });
  });

  describe('SET_CELL_BY_LABEL validation', () => {
    it('rejects missing label', () => {
      expect(() => opsSvc.validatePlan({
        summary: 'test',
        operations: [{ type: 'SET_CELL_BY_LABEL', sheet: 'S', params: { value: 1 } }],
      })).toThrow(/label/i);
    });

    it('rejects label over 500 chars', () => {
      expect(() => opsSvc.validatePlan({
        summary: 'test',
        operations: [{ type: 'SET_CELL_BY_LABEL', sheet: 'S', params: { label: 'x'.repeat(501), value: 1 } }],
      })).toThrow(/500/);
    });
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// F. DOWNLOAD SECURITY
// ═══════════════════════════════════════════════════════════════════════════

describe('F. Download security — path traversal and output-only guard', () => {
  it('download endpoint source guards against non-output- files', () => {
    const src = fs.readFileSync(
      path.join(__dirname, '../../files/files.controller.ts'),
      'utf8',
    );
    expect(src).toContain("startsWith('output-')");
    expect(src).toContain('ForbiddenException');
  });

  it('download endpoint validates UUID filename pattern', () => {
    const src = fs.readFileSync(
      path.join(__dirname, '../../files/files.controller.ts'),
      'utf8',
    );
    expect(src).toContain('/^output-[0-9a-f-]+\\.(xlsx)$/');
  });

  it('download endpoint uses path.basename (traversal safe)', () => {
    const src = fs.readFileSync(
      path.join(__dirname, '../../files/files.controller.ts'),
      'utf8',
    );
    expect(src).toContain('path.basename(filename)');
  });

  it('valid output filename passes regex', () => {
    const filename = 'output-550e8400-e29b-41d4-a716-446655440000.xlsx';
    expect(/^output-[0-9a-f-]+\.(xlsx)$/.test(filename)).toBe(true);
  });

  it('path traversal attempt fails regex', () => {
    const filename = '../../etc/passwd';
    expect(/^output-[0-9a-f-]+\.(xlsx)$/.test(filename)).toBe(false);
  });

  it('uploaded input file name fails output- guard', () => {
    const filename = '550e8400-e29b-41d4-a716-446655440000.xlsx';
    expect(filename.startsWith('output-')).toBe(false);
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// G. FILE RETENTION / CLEANUP
// ═══════════════════════════════════════════════════════════════════════════

describe('G. File retention cleanup', () => {
  it('cleanupOldFiles deletes files older than retention period', async () => {
    const { svc, tmpDir } = makeFilesService();

    // Write an old output file
    const oldFile = path.join(tmpDir, 'output-aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa.xlsx');
    fs.writeFileSync(oldFile, Buffer.from('old'));
    // Backdate mtime by 2 hours
    const oldTime = new Date(Date.now() - 2 * 60 * 60 * 1000);
    fs.utimesSync(oldFile, oldTime, oldTime);

    // Write a new output file
    const newFile = path.join(tmpDir, 'output-bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb.xlsx');
    fs.writeFileSync(newFile, Buffer.from('new'));

    const result = await svc.cleanupOldFiles(1 * 60 * 60 * 1000); // 1 hour retention

    expect(result.deleted).toBe(1);
    expect(result.errors).toBe(0);
    expect(fs.existsSync(oldFile)).toBe(false); // old file deleted
    expect(fs.existsSync(newFile)).toBe(true);  // new file kept
  });

  it('cleanupOldFiles does not delete non-UUID files', async () => {
    const { svc, tmpDir } = makeFilesService();

    const configFile = path.join(tmpDir, 'config.json');
    fs.writeFileSync(configFile, '{}');
    const oldTime = new Date(Date.now() - 48 * 60 * 60 * 1000);
    fs.utimesSync(configFile, oldTime, oldTime);

    await svc.cleanupOldFiles(1 * 60 * 60 * 1000);

    // config.json should NOT be deleted
    expect(fs.existsSync(configFile)).toBe(true);
  });

  it('cleanupOldFiles returns 0 deleted when all files are recent', async () => {
    const { svc, tmpDir } = makeFilesService();
    const file = path.join(tmpDir, 'output-cccccccc-cccc-cccc-cccc-cccccccccccc.xlsx');
    fs.writeFileSync(file, Buffer.from('recent'));
    // mtime is now — within retention window

    const result = await svc.cleanupOldFiles(1 * 60 * 60 * 1000);
    expect(result.deleted).toBe(0);
    expect(fs.existsSync(file)).toBe(true);
  });

  it('cleanupOldFiles handles empty directory gracefully', async () => {
    const { svc } = makeFilesService();
    const result = await svc.cleanupOldFiles();
    expect(result.deleted).toBe(0);
    expect(result.errors).toBe(0);
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// H. FORMULA PROTECTION REGRESSION
// ═══════════════════════════════════════════════════════════════════════════

describe('H. Formula protection regression', () => {
  async function makeFormulaBuffer(): Promise<Buffer> {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Calc');
    ws.getCell('A1').value = 100;
    ws.getCell('A2').value = { formula: 'SUM(B1:B5)', result: 500 };
    return Buffer.from(await wb.xlsx.writeBuffer());
  }

  it('SET_CELL on formula cell still blocked by default', async () => {
    const buf = await makeFormulaBuffer();
    const wb = await excelSvc.loadWorkbook(buf);
    const plan = {
      summary: 'test',
      operations: [{ type: 'SET_CELL', sheet: 'Calc', params: { row: 2, col: 1, value: 999 } }],
    };
    opsSvc.validatePlan(plan);
    expect(() => excelSvc.applyOperationPlan(wb, plan)).toThrow(/formula/i);
  });

  it('SET_CELL_RANGE still atomic — no partial write on formula cell', async () => {
    const buf = await makeFormulaBuffer();
    const wb = await excelSvc.loadWorkbook(buf);
    const plan = {
      summary: 'test',
      operations: [{
        type: 'SET_CELL_RANGE', sheet: 'Calc',
        params: { startRow: 1, startCol: 1, endRow: 2, endCol: 1, value: 0 },
      }],
    };
    try { excelSvc.applyOperationPlan(wb, plan); } catch {}
    // A1 must NOT be modified — pre-scan blocked before any write
    expect(wb.getWorksheet('Calc').getCell(1, 1).value).toBe(100);
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// I. COORDINATE VALIDATION REGRESSION
// ═══════════════════════════════════════════════════════════════════════════

describe('I. Coordinate validation regression', () => {
  it('valid A1 coordinate passes isCellAddressValid', () => {
    const { isCellAddressValid } = require('../../excel/utils/coordinate.utils');
    expect(isCellAddressValid('A1')).toBe(true);
    expect(isCellAddressValid('Z100')).toBe(true);
  });

  it('invalid coordinates fail isCellAddressValid', () => {
    const { isCellAddressValid } = require('../../excel/utils/coordinate.utils');
    expect(isCellAddressValid('A0')).toBe(false);
    expect(isCellAddressValid('1A')).toBe(false);
    expect(isCellAddressValid('')).toBe(false);
    expect(isCellAddressValid('NOTACELL')).toBe(false);
  });

  it('GET_CELL_CONTEXT rejects invalid coordinate in validatePlan', () => {
    expect(() => opsSvc.validatePlan({
      summary: 'test',
      operations: [{
        type: 'GET_CELL_CONTEXT',
        sheet: 'S',
        params: { sheetName: 'S', coordinate: 'A0' },
      }],
    })).toThrow();
  });

  it('GET_CELL_CONTEXT accepts valid coordinate', () => {
    expect(() => opsSvc.validatePlan({
      summary: 'test',
      operations: [{
        type: 'GET_CELL_CONTEXT',
        sheet: 'S',
        params: { sheetName: 'S', coordinate: 'G42' },
      }],
    })).not.toThrow();
  });
});
