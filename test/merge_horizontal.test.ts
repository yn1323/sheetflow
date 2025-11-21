import { describe, it, expect } from 'vitest';
import { createWorkbook, defineSheet } from '../src';
import { readExcel } from './utils';
import * as path from 'path';
import * as fs from 'fs';

const OUTPUT_DIR = path.join(__dirname, 'output');
if (!fs.existsSync(OUTPUT_DIR)) {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

describe('Horizontal Merge', () => {
  it('should merge horizontally adjacent identical values', async () => {
    const filePath = path.join(OUTPUT_DIR, 'merge_horizontal.xlsx');
    interface Data { a: string; b: string; c: string; d: string }
    const data: Data[] = [
        { a: 'X', b: 'X', c: 'Y', d: 'Z' }, // X-X merge
        { a: 'A', b: 'B', c: 'C', d: 'C' }, // C-C merge (cols 3-4)
        { a: 'P', b: 'P', c: 'P', d: 'Q' }  // P-P-P merge
    ];

    const sheetDef = defineSheet<Data>({
      name: 'HorizontalMerge',
      columns: [
        { key: 'a', header: 'A', merge: 'horizontal' },
        { key: 'b', header: 'B', merge: 'horizontal' },
        { key: 'c', header: 'C', merge: 'horizontal' },
        { key: 'd', header: 'D', merge: 'horizontal' }
      ]
    });

    await createWorkbook().addSheet(sheetDef, data).save(filePath);

    const workbook = await readExcel(filePath);
    const sheet = workbook.getWorksheet('HorizontalMerge');
    
    if(sheet) {
        // Row 2 (Data 1): X, X, Y, Z -> Merge(1,2), 3, 4
        const r2c1 = sheet.getCell(2, 1);
        expect(r2c1.value).toBe('X');
        expect(r2c1.isMerged).toBe(true);
        expect(sheet.getCell(2, 2).master).toBe(r2c1);
        expect(sheet.getCell(2, 3).isMerged).toBe(false);

        // Row 3 (Data 2): A, B, C, C -> 1, 2, Merge(3,4)
        expect(sheet.getCell(3, 1).isMerged).toBe(false);
        const r3c3 = sheet.getCell(3, 3);
        expect(r3c3.value).toBe('C');
        expect(r3c3.isMerged).toBe(true);
        expect(sheet.getCell(3, 4).master).toBe(r3c3);

        // Row 4 (Data 3): P, P, P, Q -> Merge(1,2,3), 4
        const r4c1 = sheet.getCell(4, 1);
        expect(r4c1.value).toBe('P');
        expect(r4c1.isMerged).toBe(true);
        expect(sheet.getCell(4, 2).master).toBe(r4c1);
        expect(sheet.getCell(4, 3).master).toBe(r4c1);
        expect(sheet.getCell(4, 4).isMerged).toBe(false);
    }
  });
});
