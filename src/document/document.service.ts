import { Injectable } from '@nestjs/common';
import * as PizZip from 'pizzip';
import * as Docxtemplater from 'docxtemplater';
import { Transform, Writable } from 'stream';

@Injectable()
export class DocumentService {
  // Чтение файла Word и поиск переменных
  async extractVariables(buffer: Buffer): Promise<string[]> {
    const zip = new PizZip(buffer);

    const doc = new Docxtemplater(zip, {
      delimiters: { start: '<~!', end: '!~>' },
    });
    const content = doc.getFullText();
    const variables = content.match(/<~!.*?!~>/g) || [];
    return variables.map((v) => v.replace(/<~!|!~>/g, ''));
  }

  // Замена переменных в документе
  async replaceVariables(
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
