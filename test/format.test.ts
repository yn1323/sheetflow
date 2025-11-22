import { describe, it, expect } from 'vitest';
import { createWorkbook } from '../src';
import { readExcel } from './utils';
import * as path from 'path';
import * as fs from 'fs';
import ExcelJS from 'exceljs';

const OUTPUT_DIR = path.join(__dirname, 'output');
if (!fs.existsSync(OUTPUT_DIR)) {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

describe('Formatting', () => {
  it('should apply number and date formats', async () => {
    const filePath = path.join(OUTPUT_DIR, 'format.xlsx');

    await createWorkbook().addSheet({
      name: 'Format',
      headers: [
        { key: 'date', label: 'Date', format: 'yyyy-mm-dd' },
        { key: 'money', label: 'Money', format: '$#,##0' }
      ],
      rows: [{ date: new Date('2025-01-01'), money: 1000 }]
    }).save(filePath);

    const workbook = await readExcel(filePath);
    const sheet = workbook.getWorksheet('Format');
    
    if(sheet) {
      expect(sheet.getCell(2, 1).numFmt).toBe('yyyy-mm-dd');
      expect(sheet.getCell(2, 2).numFmt).toBe('$#,##0');
    }
  });

  it('should preserve leading zeros in strings', async () => {
    const filePath = path.join(OUTPUT_DIR, 'zeros.xlsx');

    await createWorkbook().addSheet({
      name: 'Zeros',
      headers: [{ key: 'code', label: 'Code' }],
      rows: [{ code: '007' }]
    }).save(filePath);

    const workbook = await readExcel(filePath);
    const sheet = workbook.getWorksheet('Zeros');
    
    if(sheet) {
      const cell = sheet.getCell(2, 1);
      expect(cell.value).toBe('007');
      expect(cell.type).toBe(ExcelJS.ValueType.String);
    }
  });
});
