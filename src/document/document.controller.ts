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
import * as PizZip from 'pizzip';
import * as Docxtemplater from 'docxtemplater';
import { Transform, Writable } from 'stream';

@Controller('document')
export class DocumentController {
  constructor() {}

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

    const variables = await this.extractVariablesFromBuffer(file.buffer);
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

    const updatedBuffer = await this.replaceVariablesFromBuffer(
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

  // Чтение файла Word и поиск переменных
  async extractVariablesFromBuffer(buffer: Buffer): Promise<string[]> {
    const zip = new PizZip(buffer);

    const doc = new Docxtemplater(zip, {
      delimiters: { start: '<~!', end: '!~>' },
    });
    const content = doc.getFullText();
    const variables = content.match(/<~!.*?!~>/g) || [];
    return variables.map((v) => v.replace(/<~!|!~>/g, ''));
  }

  // Замена переменных в документе
  async replaceVariablesFromBuffer(
    buffer: Buffer,
    replacements: Record<string, string | string[]>,
  ): Promise<Buffer> {
    // Счётчики для отслеживания использования каждого тега
    const usageCounters: Record<string, number> = {};

    // Создаём поток чтения и обработки
    const transformStream = new Transform({
      transform(chunk, encoding, callback) {
        let modifiedBuffer: Buffer;

        try {
          const zip = new PizZip(chunk);

          const doc = new Docxtemplater(zip, {
            delimiters: { start: '<~!', end: '!~>' },
            parser: (tag) => {
              // Используем регулярное выражение с поддержкой Unicode для поиска переменных
              const unicodeRegex = /^[\p{L}\p{N}_]+$/u; // Учитывает буквы, цифры и подчёркивание
              if (!unicodeRegex.test(tag)) {
                throw new Error(
                  `Invalid tag format: ${tag}. Only letters, numbers, and underscores are allowed.`,
                );
              }

              if (usageCounters[tag] === undefined) {
                usageCounters[tag] = 0;
              }

              return {
                get: () => {
                  const value = replacements[tag];
                  if (Array.isArray(value)) {
                    const currentIndex = usageCounters[tag]++;
                    return value[currentIndex] !== undefined
                      ? value[currentIndex]
                      : `<~!${tag}!~>`;
                  }
                  return value !== undefined ? value : `<~!${tag}!~>`;
                },
              };
            },
          });

          doc.render();

          // Получаем модифицированный документ как Buffer
          modifiedBuffer = Buffer.from(
            doc.getZip().generate({ type: 'nodebuffer' }),
          );
        } catch (err) {
          return callback(err);
        }

        // Передаём модифицированный кусок дальше в поток
        callback(null, modifiedBuffer);
      },
    });

    // Создаём поток вывода для сохранения результата
    const output: Writable = new Writable();

    // Сохраняем финальный результат в переменной
    const resultChunks: Buffer[] = [];
    output._write = (chunk, encoding, callback) => {
      resultChunks.push(chunk);
      callback();
    };

    return new Promise<Buffer>((resolve, reject) => {
      // Слушаем завершение записи
      output.on('finish', () => {
        resolve(Buffer.concat(resultChunks));
      });

      // Обрабатываем ошибки
      transformStream.on('error', (err) => {
        reject(err);
      });

      // Подключаем буфер к потокам
      transformStream.pipe(output);

      // Запускаем поток с исходным буфером
      transformStream.write(buffer);
      transformStream.end();
    });
  }
}
