import { Module } from '@nestjs/common';
import { AiService } from './ai.service';
import { AiController } from './ai.controller';
import { OperationsModule } from '../operations/operations.module';

@Module({
  imports: [
    OperationsModule, // provides OperationsService for plan validation
  ],
  controllers: [AiController],
  providers: [AiService],
  exports: [AiService],
})
export class AiModule {}
