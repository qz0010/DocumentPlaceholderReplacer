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
    console.log('replacements', JSON.stringify(replacements));

    const decodeIfBroken = (value: string): string => {
      try {
        const decoded = Buffer.from(value, 'latin1').toString('utf-8');
        // If the decoded string has more valid Cyrillic characters, use it
        const originalCyrillicCount = (value.match(/[\u0400-\u04FF]/g) || [])
          .length;
        const decodedCyrillicCount = (decoded.match(/[\u0400-\u04FF]/g) || [])
          .length;

        return decodedCyrillicCount > originalCyrillicCount ? decoded : value;
      } catch {
        return value; // Return original if decoding fails
      }
    };

    // Decode replacements
    const decodedReplacements = Object.fromEntries(
      Object.entries(replacements).map(([key, value]) => [
        decodeIfBroken(key), // Decode key if broken
        Array.isArray(value)
          ? value.map((v) => decodeIfBroken(v)) // Decode array values
          : decodeIfBroken(value), // Decode single value
      ]),
    );

    console.log('decodedReplacements', JSON.stringify(decodedReplacements));

    const updatedBuffer = await this.documentService.replaceVariables(
      file.buffer,
      decodedReplacements,
    );

    res.set({
      'Content-Type':
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'Content-Disposition': 'attachment; filename=updated-document.docx',
    });
    res.send(updatedBuffer);
  }
}
