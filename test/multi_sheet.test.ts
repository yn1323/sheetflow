import { describe, it, expect } from 'vitest';
import { createWorkbook } from '../src';
import { readExcel } from './utils';
import * as path from 'path';
import * as fs from 'fs';

const OUTPUT_DIR = path.join(__dirname, 'output');
if (!fs.existsSync(OUTPUT_DIR)) {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

describe('Multiple Sheets', () => {
  it('should create a workbook with multiple sheets', async () => {
    const filePath = path.join(OUTPUT_DIR, 'multi.xlsx');

    await createWorkbook()
      .addSheet({
        name: 'Users',
        headers: [{ key: 'name', label: 'Name' }],
        rows: [{ name: 'Alice' }]
      })
      .addSheet({
        name: 'Products',
        headers: [{ key: 'title', label: 'Title' }],
        rows: [{ title: 'Laptop' }]
      })
      .save(filePath);

    const workbook = await readExcel(filePath);
    
    const sheet1 = workbook.getWorksheet('Users');
    const sheet2 = workbook.getWorksheet('Products');
    
    expect(sheet1).toBeDefined();
    expect(sheet2).toBeDefined();
    
    if(sheet1 && sheet2) {
      expect(sheet1.getCell(2, 1).value).toBe('Alice');
      expect(sheet2.getCell(2, 1).value).toBe('Laptop');
    }
  });
});
