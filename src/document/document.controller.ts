import {
  Controller,
  Post,
  UploadedFile,
  Body,
  UseInterceptors,
  BadRequestException,
  Res,
} from '@nestjs/common';
import { FileInterceptor } from '@nestjs/platform-express';
import { DocumentService } from './document.service';
import { Response } from 'express';

@Controller('document')
export class DocumentController {
  constructor(private readonly documentService: DocumentService) {}

  // 1. Извлечение переменных из Word-документа
  @Post('extract')
  @UseInterceptors(FileInterceptor('file'))
  async extractVariables(@UploadedFile() file: Express.Multer.File) {
    if (
      !file ||
      file.mimetype !==
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    ) {
      throw new BadRequestException(
        'Invalid file format. Please upload a .docx file.',
      );
    }

    const variables = await this.documentService.extractVariables(file.buffer);
    return { variables };
  }

  // 2. Замена переменных в Word-документе
  @Post('replace')
  @UseInterceptors(FileInterceptor('file'))
  async replaceVariables(
    @UploadedFile() file: Express.Multer.File,
    @Body() replacements: Record<string, string | string[]>,
    @Res() res: Response,
  ) {
    if (
      !file ||
      file.mimetype !==
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    ) {
      throw new BadRequestException(
        'Invalid file format. Please upload a .docx file.',
      );
    }

    const updatedBuffer = await this.documentService.replaceVariables(
      file.buffer,
      replacements,
    );

    res.set({
      'Content-Type':
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'Content-Disposition': 'attachment; filename=updated-document.docx',
    });
    res.send(updatedBuffer);
  }
}
