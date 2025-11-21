import { describe, it, expect } from 'vitest';
import { createWorkbook, defineSheet } from '../src';
import { readExcel } from './utils';
import * as path from 'path';
import * as fs from 'fs';

const OUTPUT_DIR = path.join(__dirname, 'output');
if (!fs.existsSync(OUTPUT_DIR)) {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

describe('Complex Header', () => {
  it('should render multi-line headers with merges', async () => {
    const filePath = path.join(OUTPUT_DIR, 'header_complex.xlsx');
    interface Data { col1: string; col2: string; col3: string }
    const data: Data[] = [{ col1: 'A', col2: 'B', col3: 'C' }];

    const sheetDef = defineSheet<Data>({
      name: 'ComplexHeader',
      columns: [
        { key: 'col1', header: 'Col 1' },
        { key: 'col2', header: 'Col 2' },
        { key: 'col3', header: 'Col 3' }
      ],
      header: {
        rows: [
          [
            { value: 'Group 1', colSpan: 2, style: { alignment: { horizontal: 'center' } } },
            { value: 'Group 2', rowSpan: 2 }
          ],
          [
            'Sub 1', 'Sub 2'
          ]
        ]
      }
    });

    await createWorkbook().addSheet(sheetDef, data).save(filePath);

    const workbook = await readExcel(filePath);
    const sheet = workbook.getWorksheet('ComplexHeader');
    
    if(sheet) {
        // Check Group 1 merge
        const group1 = sheet.getCell(1, 1);
        expect(group1.value).toBe('Group 1');
        expect(group1.isMerged).toBe(true);
        // Check range of merge? ExcelJS doesn't make it super easy to check exact range without iterating master.
        // But we can check if cell(1,2) is merged to (1,1)
        const cell12 = sheet.getCell(1, 2);
        expect(cell12.master).toBe(group1);

        // Check Group 2 merge (rowSpan)
        const group2 = sheet.getCell(1, 3);
        expect(group2.value).toBe('Group 2');
        expect(group2.isMerged).toBe(true);
        const cell23 = sheet.getCell(2, 3);
        expect(cell23.master).toBe(group2);

        // Check Sub headers
        expect(sheet.getCell(2, 1).value).toBe('Sub 1');
        expect(sheet.getCell(2, 2).value).toBe('Sub 2');

        // Check Data start row (should be 3)
        expect(sheet.getCell(3, 1).value).toBe('A');
    }
  });
});
