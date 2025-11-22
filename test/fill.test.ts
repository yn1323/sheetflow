import { describe, it, expect } from 'vitest';
import { createWorkbook } from '../src';
import { readExcel, getCellStyle } from './utils';
import * as path from 'path';
import * as fs from 'fs';

const OUTPUT_DIR = path.join(__dirname, 'output');
if (!fs.existsSync(OUTPUT_DIR)) {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

describe('Fill Styling', () => {
  it('should apply fill styles to headers and rows', async () => {
    const filePath = path.join(OUTPUT_DIR, 'fill.xlsx');

    await createWorkbook().addSheet({
      name: 'FillTest',
      title: {
        label: 'ID List',
        style: { fill: { color: '#CCCCCC' } }
      },
      headers: [
        { key: 'id', label: 'ID' }
      ],
      rows: [
        { id: 1 },
        { id: 2, styles: { fill: { color: '#EFEFEF' } } },
        { id: 3 },
      ],
      styles: {
        row: (_, index) => index % 2 === 1 ? { fill: { color: '#EFEFEF' } } : {}
      }
    }).save(filePath);

    const workbook = await readExcel(filePath);
    const sheet = workbook.getWorksheet('FillTest');
    
    expect(sheet).toBeDefined();
    if(sheet) {
        // Title Fill
        const titleStyle = getCellStyle(sheet, 1, 1);
        expect(titleStyle.fill).toMatchObject({
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFCCCCCC' }
        });

        // Row Fill (index 1 in data -> Excel Row 3 because of title + header)
        // data[0] (id:1) -> index 0 -> no fill -> Excel Row 3
        // data[1] (id:2) -> index 1 -> fill -> Excel Row 4
        
        const row3Style = getCellStyle(sheet, 3, 1);
        expect(row3Style.fill).toBeUndefined();

        const row4Style = getCellStyle(sheet, 4, 1);
        expect(row4Style.fill).toMatchObject({
             type: 'pattern',
             pattern: 'solid',
             fgColor: { argb: 'FFEFEFEF' }
        });
    }
  });
});
