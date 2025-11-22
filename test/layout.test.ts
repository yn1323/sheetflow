import { describe, it, expect } from 'vitest';
import { createWorkbook } from '../src';
import { readExcel } from './utils';
import * as path from 'path';
import * as fs from 'fs';

const OUTPUT_DIR = path.join(__dirname, 'output');
if (!fs.existsSync(OUTPUT_DIR)) {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

describe('Layout Features', () => {
  it('should handle vertical merge', async () => {
    const filePath = path.join(OUTPUT_DIR, 'merge.xlsx');

    await createWorkbook().addSheet({
      name: 'Merge',
      headers: [
        { key: 'group', label: 'Group', merge: 'vertical' },
        { key: 'item', label: 'Item' }
      ],
      rows: [
        { group: 'A', item: '1' },
        { group: 'A', item: '2' },
        { group: 'B', item: '3' }
      ]
    }).save(filePath);

    const workbook = await readExcel(filePath);
    const sheet = workbook.getWorksheet('Merge');
    
    if(sheet) {
      const cell1 = sheet.getCell(2, 1);
      const cell2 = sheet.getCell(3, 1);
      expect(cell1.value).toBe('A');
      expect(cell2.master).toBe(cell1);
    }
  });

  it('should calculate auto width', async () => {
    const filePath = path.join(OUTPUT_DIR, 'width.xlsx');

    await createWorkbook().addSheet({
      name: 'Width',
      headers: [{ key: 'text', label: 'Text', width: 'auto' }],
      rows: [{ text: 'This is a very long text to test auto width' }]
    }).save(filePath);

    const workbook = await readExcel(filePath);
    const sheet = workbook.getWorksheet('Width');
    
    if(sheet) {
      const col = sheet.getColumn(1);
      expect(col.width).toBeGreaterThan(20);
    }
  });
});
