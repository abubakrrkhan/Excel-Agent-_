import {
  Injectable,
  Logger,
  BadRequestException,
  NotFoundException,
  InternalServerErrorException,
} from '@nestjs/common';
import * as path from 'path';
import * as fs from 'fs';
import * as crypto from 'crypto';
import { DatabaseService } from '../database/database.service';

const ALLOWED_MIME_TYPES = [
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
  // NOTE: .xls (application/vnd.ms-excel) is intentionally excluded.
  // ExcelJS can read .xls but always writes .xlsx. Accepting .xls would
  // produce an .xlsx output with a mismatched filename, which confuses users.
  // Reject .xls explicitly and ask users to convert to .xlsx first.
];

const ALLOWED_EXTENSIONS = ['.xlsx'];

const MAX_FILE_SIZE_BYTES = 10 * 1024 * 1024; // 10MB

// ─── Shape of a persisted ExcelFile record ────────────────────────────────────

export interface ExcelFileRecord {
  id: string;
  originalName: string;
  storedName: string;
  mimeType: string;
  sizeBytes: number;
  uploadedAt: Date;
  updatedAt: Date;
  status: string;
}

// ─── Service ──────────────────────────────────────────────────────────────────

@Injectable()
export class FilesService {
  private readonly logger = new Logger(FilesService.name);
  private readonly uploadDir = path.join(process.cwd(), 'uploads');

  constructor(private readonly databaseService: DatabaseService) {
    this.ensureUploadDir();
  }

  // ── Validation ──────────────────────────────────────────────────────────────

  validateExcelFile(file: Express.Multer.File): void {
    if (!file) {
      throw new BadRequestException('No file provided');
    }

    // Check extension first — gives the clearest error for .xls uploads
    const ext = path.extname(file.originalname).toLowerCase();
    if (ext === '.xls') {
      throw new BadRequestException(
        '.xls format is not supported. Please convert your file to .xlsx (Excel 2007+) and try again.',
      );
    }
    if (!ALLOWED_EXTENSIONS.includes(ext)) {
      throw new BadRequestException(
        `Unsupported file extension "${ext}". Only .xlsx files are accepted.`,
      );
    }

    if (!ALLOWED_MIME_TYPES.includes(file.mimetype)) {
      throw new BadRequestException(
        'Invalid file type. Only .xlsx Excel files are allowed.',
      );
    }
    if (file.size > MAX_FILE_SIZE_BYTES) {
      throw new BadRequestException(
        `File too large. Maximum size is ${MAX_FILE_SIZE_BYTES / 1024 / 1024} MB.`,
      );
    }
  }

  // ── Disk I/O ────────────────────────────────────────────────────────────────

  /**
   * Write an uploaded file buffer to the uploads directory.
   * Returns the stored filename (uuid + original extension).
   */
  async saveFileToDisk(file: Express.Multer.File): Promise<string> {
    const ext = path.extname(file.originalname).toLowerCase() || '.xlsx';
    const storedName = `${crypto.randomUUID()}${ext}`;
    const filePath = path.join(this.uploadDir, storedName);

    try {
      await fs.promises.writeFile(filePath, file.buffer);
      this.logger.log(`Saved file to disk: ${storedName} (${file.size} bytes)`);
      return storedName;
    } catch (err) {
      this.logger.error('Failed to write file to disk', err.message);
      throw new InternalServerErrorException('Could not save uploaded file');
    }
  }

  /**
   * Read a previously saved file from the uploads directory into a Buffer.
   * Applies basename sanitisation — no directory traversal possible.
   */
  async readFileFromDisk(storedName: string): Promise<Buffer> {
    const safeName = path.basename(storedName);
    const filePath = path.join(this.uploadDir, safeName);

    try {
      return await fs.promises.readFile(filePath);
    } catch (err) {
      this.logger.error(`Failed to read file: ${safeName}`, err.message);
      throw new NotFoundException(`File not found on disk: ${safeName}`);
    }
  }

  /**
   * Save a processed output Buffer to the uploads directory.
   * Always generates a fresh UUID — NEVER overwrites the original file.
   * Returns the stored output filename.
   */
  async saveOutputFile(buffer: Buffer, originalName: string): Promise<string> {
    const ext = path.extname(originalName).toLowerCase() || '.xlsx';
    const outputName = `output-${crypto.randomUUID()}${ext}`;
    const filePath = path.join(this.uploadDir, outputName);

    try {
      await fs.promises.writeFile(filePath, buffer);
      this.logger.log(`Saved output file: ${outputName} (${buffer.length} bytes)`);
      return outputName;
    } catch (err) {
      this.logger.error('Failed to save output file', err.message);
      throw new InternalServerErrorException('Could not save processed Excel file');
    }
  }

  getUploadPath(filename: string): string {
    return path.join(this.uploadDir, path.basename(filename));
  }

  /**
   * Resolve a stored filename to its absolute disk path (basename-safe).
   */
  getOutputFilePath(storedName: string): string {
    return path.join(this.uploadDir, path.basename(storedName));
  }

  /**
   * Check whether a file exists on disk (basename-safe).
   */
  fileExistsOnDisk(storedName: string): boolean {
    return fs.existsSync(path.join(this.uploadDir, path.basename(storedName)));
  }

  // ── Database ────────────────────────────────────────────────────────────────

  /**
   * Insert an ExcelFile record into the database.
   * Falls back gracefully when DATABASE_URL is not configured.
   */
  async createFileRecord(params: {
    id: string;
    originalName: string;
    storedName: string;
    mimeType: string;
    sizeBytes: number;
  }): Promise<ExcelFileRecord | null> {
    const pool = this.databaseService.getPool();

    if (!pool) {
      this.logger.warn('No DB pool — skipping ExcelFile record creation');
      return null;
    }

    try {
      const result = await pool.query<ExcelFileRecord>(
        `INSERT INTO excel_files (id, original_name, stored_name, mime_type, size_bytes, status, uploaded_at, updated_at)
         VALUES ($1, $2, $3, $4, $5, 'PENDING', NOW(), NOW())
         RETURNING id, original_name AS "originalName", stored_name AS "storedName",
                   mime_type AS "mimeType", size_bytes AS "sizeBytes",
                   uploaded_at AS "uploadedAt", updated_at AS "updatedAt", status`,
        [params.id, params.originalName, params.storedName, params.mimeType, params.sizeBytes],
      );

      this.logger.log(`Created ExcelFile record: ${params.id}`);
      return result.rows[0];
    } catch (err) {
      this.logger.error('Failed to create ExcelFile record', err.message);
      throw new InternalServerErrorException('Could not save file metadata to database');
    }
  }

  /**
   * Fetch an ExcelFile record by ID.
   */
  async findFileRecord(id: string): Promise<ExcelFileRecord> {
    const pool = this.databaseService.getPool();

    if (!pool) {
      throw new InternalServerErrorException('Database is not available');
    }

    try {
      const result = await pool.query<ExcelFileRecord>(
        `SELECT id, original_name AS "originalName", stored_name AS "storedName",
                mime_type AS "mimeType", size_bytes AS "sizeBytes",
                uploaded_at AS "uploadedAt", updated_at AS "updatedAt", status
         FROM excel_files WHERE id = $1`,
        [id],
      );

      if (result.rows.length === 0) {
        throw new NotFoundException(`File with ID "${id}" not found`);
      }

      return result.rows[0];
    } catch (err) {
      if (err instanceof NotFoundException) throw err;
      this.logger.error(`Failed to fetch ExcelFile record ${id}`, err.message);
      throw new InternalServerErrorException('Could not retrieve file metadata');
    }
  }

  // ── Retention / cleanup ─────────────────────────────────────────────────────

  /**
   * Delete files older than retentionMs from the uploads directory.
   *
   * @param retentionMs    Files modified before (now - retentionMs) are deleted.
   * @param fileType       'input' | 'output' | 'all'
   * @param protectedFiles Set of filenames that must not be deleted.
   *                       CleanupService passes filenames of queued/active jobs.
   *
   * Safety rules:
   *   - Only UUID-named files matching expected patterns are touched.
   *   - Files in protectedFiles are skipped regardless of age.
   *   - Individual file errors are counted but never crash the whole pass.
   */
  async cleanupOldFiles(
    retentionMs: number = 24 * 60 * 60 * 1000,
    fileType: 'input' | 'output' | 'all' = 'all',
    protectedFiles: Set<string> = new Set(),
  ): Promise<{ deleted: number; errors: number }> {
    let deleted = 0;
    let errors = 0;
    const cutoff = Date.now() - retentionMs;

    const inputPattern = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\.xlsx$/;
    const outputPattern = /^output-[0-9a-f-]+\.xlsx$/;

    try {
      const files = await fs.promises.readdir(this.uploadDir);

      for (const file of files) {
        const isInput = inputPattern.test(file);
        const isOutput = outputPattern.test(file);

        if (fileType === 'input' && !isInput) continue;
        if (fileType === 'output' && !isOutput) continue;
        if (fileType === 'all' && !isInput && !isOutput) continue;

        // Never delete a file currently referenced by a queued or active job
        if (protectedFiles.has(file)) {
          this.logger.debug(`Cleanup: skipping protected file ${file}`);
          continue;
        }

        const filePath = path.join(this.uploadDir, file);

        try {
          const stat = await fs.promises.stat(filePath);
          if (stat.mtimeMs < cutoff) {
            await fs.promises.unlink(filePath);
            deleted++;
            this.logger.debug(`Cleanup: deleted ${file}`);
          }
        } catch (err) {
          errors++;
          this.logger.warn(`Cleanup: could not process ${file}: ${err.message}`);
        }
      }
    } catch (err) {
      this.logger.error(`Cleanup: failed to read upload directory: ${err.message}`);
      errors++;
    }

    this.logger.log(`Cleanup [${fileType}]: ${deleted} deleted, ${errors} errors, ${protectedFiles.size} protected`);
    return { deleted, errors };
  }

  private ensureUploadDir() {
    if (!fs.existsSync(this.uploadDir)) {
      fs.mkdirSync(this.uploadDir, { recursive: true });
      this.logger.log(`Upload directory created: ${this.uploadDir}`);
    }
  }
}
