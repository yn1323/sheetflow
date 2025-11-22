import { describe, it, expect } from 'vitest';
import { createWorkbook } from '../src';
import { readExcel, getCellStyle } from './utils';
import * as path from 'path';
import * as fs from 'fs';

const OUTPUT_DIR = path.join(__dirname, 'output');
if (!fs.existsSync(OUTPUT_DIR)) {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

describe('Font Styling', () => {
  it('should apply font styles to columns and conditionally', async () => {
    const filePath = path.join(OUTPUT_DIR, 'font.xlsx');

    await createWorkbook().addSheet({
      name: 'FontTest',
      headers: [
        { 
          key: 'name', 
          label: 'Name', 
          style: { font: { bold: true, size: 14 } } 
        },
        { 
          key: 'status', 
          label: 'Status', 
          style: (val) => ({ 
            font: { color: val === 'active' ? '#00FF00' : '#FF0000', italic: true } 
          })
        }
      ],
      rows: [
        { name: 'Alice', status: 'active' },
        { name: 'Bob', status: 'inactive' },
      ]
    }).save(filePath);

    const workbook = await readExcel(filePath);
    const sheet = workbook.getWorksheet('FontTest');
    
    expect(sheet).toBeDefined();
    if(sheet) {
      // Column Style
      const nameStyle = getCellStyle(sheet, 2, 1);
      expect(nameStyle.font).toMatchObject({ bold: true, size: 14 });

      // Conditional Style (Active)
      const activeStyle = getCellStyle(sheet, 2, 2);
      expect(activeStyle.font?.color).toMatchObject({ argb: 'FF00FF00' });
      expect(activeStyle.font?.italic).toBe(true);

      // Conditional Style (Inactive)
      const inactiveStyle = getCellStyle(sheet, 3, 2);
      expect(inactiveStyle.font?.color).toMatchObject({ argb: 'FFFF0000' });
      expect(inactiveStyle.font?.italic).toBe(true);
    }
  });
});
