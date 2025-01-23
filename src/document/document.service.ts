import { Injectable } from '@nestjs/common';
import * as PizZip from 'pizzip';
import * as Docxtemplater from 'docxtemplater';

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
    const zip = new PizZip(buffer);

    // Счётчики вхождений для каждого тега
    const usageCounters: Record<string, number> = {};

    const doc = new Docxtemplater(zip, {
      delimiters: { start: '<~!', end: '!~>' },
      parser: (tag) => {
        // Инициализируем счётчик для тега, если он ещё не создан
        if (usageCounters[tag] === undefined) {
          usageCounters[tag] = 0;
        }
        return {
          get: () => {
            const value = replacements[tag];
            // Если передан массив, берём элемент по счётчику
            if (Array.isArray(value)) {
              const currentIndex = usageCounters[tag]++;
              return value[currentIndex] !== undefined
                ? value[currentIndex]
                : `<~!${tag}!~>`; // Если индекс вышел за границы массива
            }
            // Если не массив, возвращаем простое значение
            return value !== undefined ? value : `<~!${tag}!~>`;
          },
        };
      },
    });

    doc.render();
    return Buffer.from(doc.getZip().generate({ type: 'nodebuffer' }));
  }

}
