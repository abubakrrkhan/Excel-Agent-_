import {
  Controller,
  Post,
  Body,
  HttpCode,
  HttpStatus,
  Logger,
} from '@nestjs/common';
import { AiService, OperationPlan } from './ai.service';
import { GeneratePlanDto } from './dto/generate-plan.dto';

@Controller('ai')
export class AiController {
  private readonly logger = new Logger(AiController.name);

  constructor(private readonly aiService: AiService) {}

  /**
   * POST /api/ai/generate-plan
   *
   * Accepts workbook context + natural-language instruction.
   * Returns a validated structured operation plan.
   *
   * The Excel file is NOT modified by this endpoint.
   * Operation execution happens in a later stage (Day 4).
   */
  @Post('generate-plan')
  @HttpCode(HttpStatus.OK)
  async generatePlan(@Body() dto: GeneratePlanDto): Promise<OperationPlan> {
    this.logger.log(
      `generate-plan request | instruction="${dto.instruction.slice(0, 80)}" ` +
        `| sheets=${dto.workbook.sheetCount}`,
    );

    return this.aiService.generateOperationPlan(
      dto.instruction,
      dto.workbook,
    );
  }
}
