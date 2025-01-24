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
import { Response } from 'express';
import * as PizZip from 'pizzip';
import * as Docxtemplater from 'docxtemplater';
import { Transform, Writable } from 'stream';
import { convert } from 'libreoffice-convert';

@Controller('document')
export class DocumentController {
  constructor() {}

  @Post('extract')
  @UseInterceptors(FileInterceptor('file'))
  async extractVariables(@UploadedFile() file: Express.Multer.File) {
    const allowedMimes = [
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document', // .docx
      'application/msword', // .doc
      'text/plain',
    ];

    if (!file || !allowedMimes.includes(file.mimetype)) {
      throw new BadRequestException(
        'Invalid file format. Please upload a .docx, .doc or .txt file.',
      );
    }

    // Если это .doc, конвертируем в .docx буфер (в памяти)
    let buffer = file.buffer;
    if (file.mimetype === 'application/msword') {
      buffer = await this.convertDocToDocx(buffer);
    }

    // Если это .txt, у нас нет структуры docx — будем извлекать плейсхолдеры напрямую из текста
    if (file.mimetype === 'text/plain') {
      return {
        variables: this.extractVariablesFromTxt(buffer),
      };
    }

    // Иначе (это уже .docx или сконвертированный .doc -> .docx) — стандартная логика
    const variables = await this.extractVariablesFromBuffer(buffer);
    return { variables };
  }

  @Post('replace')
  @UseInterceptors(FileInterceptor('file'))
  async replaceVariables(
    @UploadedFile() file: Express.Multer.File,
    @Body() replacements: Record<string, string | string[]>,
    @Res() res: Response,
  ) {
    const allowedMimes = [
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document', // .docx
      'application/msword', // .doc
      'text/plain',
    ];

    if (!file || !allowedMimes.includes(file.mimetype)) {
      throw new BadRequestException(
        'Invalid file format. Please upload a .docx, .doc or .txt file.',
      );
    }

    console.log('replacements', JSON.stringify(replacements));

    const decodeIfBroken = (value: string): string => {
      try {
        const decoded = Buffer.from(value, 'latin1').toString('utf-8');
        const originalCyrillicCount = (value.match(/[\u0400-\u04FF]/g) || [])
          .length;
        const decodedCyrillicCount = (decoded.match(/[\u0400-\u04FF]/g) || [])
          .length;
        return decodedCyrillicCount > originalCyrillicCount ? decoded : value;
      } catch {
        return value;
      }
    };

    // Проходимся по replacements и декодируем
    const decodedReplacements = Object.fromEntries(
      Object.entries(replacements).map(([key, value]) => [
        decodeIfBroken(key),
        Array.isArray(value)
          ? value.map((v) => decodeIfBroken(v))
          : decodeIfBroken(value),
      ]),
    );
    console.log('decodedReplacements', JSON.stringify(decodedReplacements));

    let buffer = file.buffer;

    // Если .doc -> сначала конвертируем в .docx
    if (file.mimetype === 'application/msword') {
      buffer = await this.convertDocToDocx(buffer);
    }

    // Если .txt -> обрабатываем напрямую как текст и возвращаем .txt
    if (file.mimetype === 'text/plain') {
      const updatedTxt = this.replaceVariablesInTxt(buffer, decodedReplacements);
      res.set({
        'Content-Type': 'text/plain; charset=utf-8',
        'Content-Disposition': 'attachment; filename=updated-document.txt',
      });
      return res.send(updatedTxt);
    }

    // Иначе (у нас уже .docx или «сконвертированный .doc -> .docx»)
    const updatedBuffer = await this.replaceVariablesFromBuffer(
      buffer,
      decodedReplacements,
    );

    res.set({
      'Content-Type':
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'Content-Disposition': 'attachment; filename=updated-document.docx',
    });
    res.send(updatedBuffer);
  }

  private async convertDocToDocx(inputBuffer: Buffer): Promise<Buffer> {
    return new Promise((resolve, reject) => {
      convert(inputBuffer, '.docx', undefined, (err, result) => {
        if (err) {
          return reject(err);
        }
        resolve(result);
      });
    });
  }

  private extractVariablesFromTxt(buffer: Buffer): string[] {
    const text = buffer.toString('utf-8');
    // Ищем все вхождения <~!...!~>
    const matches = text.match(/<~!.*?!~>/g) || [];
    // Убираем служебные символы <~! и !~>
    return matches.map((m) => m.replace(/<~!|!~>/g, ''));
  }

  /**
   * Псевдо-потоковая (chunk-based) замена плейсхолдеров в txt,
   * не меняя сигнатуру метода (по-прежнему принимает Buffer и возвращает Buffer).
   */
  private replaceVariablesInTxt(
    buffer: Buffer,
    replacements: Record<string, string | string[]>,
  ): Buffer {
    // Размер «чанка» — условно 64 KB. Можно настроить под себя.
    const CHUNK_SIZE = 64 * 1024;

    // Счётчики для случаев, когда в replacements[tag] лежит массив
    const usageCounters: Record<string, number> = {};

    // «Хвост» для незаконченных плейсхолдеров (которые могут разрываться на границе чанков)
    let leftover = '';

    // Итоговая строка, куда будем складывать результат
    let result = '';

    // Главный цикл: двигаемся по buffer с шагом CHUNK_SIZE
    for (let offset = 0; offset < buffer.length; offset += CHUNK_SIZE) {
      // Вырезаем текущий чанк
      const chunk = buffer.slice(offset, offset + CHUNK_SIZE);

      // Превращаем в строку + добавляем leftover
      let text = leftover + chunk.toString('utf-8');

      // Будем собирать «обработанную» часть (без незаконченного хвоста)
      let replacedPart = '';
      let startIndex = 0;
      let match: RegExpExecArray | null;

      // Регэксп для поиска <~!someTag!~>
      // Захватываем в группу ([\p{L}\p{N}_]+), если хотим ограничиться буквами, цифрами, подчёркиваниями
      const placeholderRegex = /<~!([\p{L}\p{N}_]+)!~>/gu;

      // Ищем все вхождения плейсхолдеров, которые полностью поместились в text
      while ((match = placeholderRegex.exec(text)) !== null) {
        const index = match.index;  // где начинается <~!tag!~>
        const fullMatch = match[0]; // сам текст <~!tag!~>
        const tagName = match[1];   // "tag" без <~! и !~>

        // Добавляем промежуточный текст, который идёт до плейсхолдера
        replacedPart += text.slice(startIndex, index);

        // Подставляем значение
        replacedPart += this.getReplacementValue(tagName, replacements, usageCounters);

        // Двигаем «указатель» дальше
        startIndex = index + fullMatch.length;
      }

      // leftover = то, что осталось после последнего полного плейсхолдера
      leftover = text.slice(startIndex);

      // Добавляем обработанную часть к итоговой строке
      result += replacedPart;
    }

    // Когда прошли все чанки, leftover может содержать незаконченный плейсхолдер (или просто хвост).
    // Но раз он незаконченный, мы не можем заменить его корректно —
    // оставим, как есть:
    result += leftover;

    // Возвращаем итоговый текст в виде Buffer
    return Buffer.from(result, 'utf-8');
  }

  /**
   * Вспомогательный метод, который возвращает нужную подстановку
   * (учитывает массивное значение replacements[tag]).
   */
  private getReplacementValue(
    tagName: string,
    replacements: Record<string, string | string[]>,
    usageCounters: Record<string, number>,
  ): string {
    // Если нет в словаре — вернём исходный плейсхолдер
    if (!(tagName in replacements)) {
      return `<~!${tagName}!~>`;
    }

    const value = replacements[tagName];
    // Если массив
    if (Array.isArray(value)) {
      if (usageCounters[tagName] === undefined) {
        usageCounters[tagName] = 0;
      }
      const currentIndex = usageCounters[tagName]++;
      // Если массив «закончился», подставляем исходный плейсхолдер
      return value[currentIndex] !== undefined
        ? value[currentIndex]
        : `<~!${tagName}!~>`;
    }

    // Если строка
    return value;
  }



  //==========================================================
  //   ОБРАБОТКА .DOCX — извлечение и замена через docxtemplater
  //==========================================================
  async extractVariablesFromBuffer(buffer: Buffer): Promise<string[]> {
    const zip = new PizZip(buffer);
    const doc = new Docxtemplater(zip, {
      delimiters: { start: '<~!', end: '!~>' },
    });
    const content = doc.getFullText();
    const variables = content.match(/<~!.*?!~>/g) || [];
    return variables.map((v) => v.replace(/<~!|!~>/g, ''));
  }

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
              // Только буквы, цифры и нижнее подчёркивание
              const unicodeRegex = /^[\p{L}\p{N}_]+$/u;
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

          modifiedBuffer = Buffer.from(
            doc.getZip().generate({ type: 'nodebuffer' }),
          );
        } catch (err) {
          return callback(err);
        }

        callback(null, modifiedBuffer);
      },
    });

    // Создаём поток вывода для сохранения результата
    const output: Writable = new Writable();
    const resultChunks: Buffer[] = [];

    output._write = (chunk, encoding, callback) => {
      resultChunks.push(chunk);
      callback();
    };

    return new Promise<Buffer>((resolve, reject) => {
      output.on('finish', () => {
        resolve(Buffer.concat(resultChunks));
      });

      transformStream.on('error', (err) => {
        reject(err);
      });

      transformStream.pipe(output);

      transformStream.write(buffer);
      transformStream.end();
    });
  }
}
