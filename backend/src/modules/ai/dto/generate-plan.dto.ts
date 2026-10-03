import {
  IsString,
  IsNotEmpty,
  MaxLength,
  IsObject,
  ValidateNested,
  IsArray,
  IsInt,
  Min,
  IsPositive,
} from 'class-validator';
import { Type } from 'class-transformer';

// ─── Nested DTOs for WorkbookInfo ─────────────────────────────────────────────

export class SheetInfoDto {
  @IsString()
  @IsNotEmpty()
  name: string;

  @IsInt()
  @Min(0)
  rowCount: number;

  @IsInt()
  @Min(0)
  dataRowCount: number;

  @IsInt()
  @Min(0)
  columnCount: number;

  @IsInt()
  @Min(0)
  dataColumnCount: number;

  @IsArray()
  headers: string[];
}

export class WorkbookInfoDto {
  @IsInt()
  @IsPositive()
  sheetCount: number;

  @IsArray()
  @ValidateNested({ each: true })
  @Type(() => SheetInfoDto)
  sheets: SheetInfoDto[];
}

// ─── Request DTO ──────────────────────────────────────────────────────────────

export class GeneratePlanDto {
  @IsString()
  @IsNotEmpty({ message: 'Instruction cannot be empty' })
  @MaxLength(2000, { message: 'Instruction must not exceed 2000 characters' })
  instruction: string;

  @IsObject()
  @ValidateNested()
  @Type(() => WorkbookInfoDto)
  workbook: WorkbookInfoDto;
}
