import { describe, it, expect } from 'vitest';
import { createWorkbook, defineSheet } from '../src';
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

    interface Data {
      id: number;
    }

    const data: Data[] = [
      { id: 1 },
      { id: 2 },
      { id: 3 },
    ];

    const sheetDef = defineSheet<Data>({
      name: 'FillTest',
      columns: [{ key: 'id', header: 'ID' }],
      header: {
          style: { fill: { color: '#CCCCCC' } }
      },
      rows: {
          style: (_, index) => index % 2 === 1 ? { fill: { color: '#EFEFEF' } } : {}
      }
    });

    await createWorkbook().addSheet(sheetDef, data).save(filePath);

    const workbook = await readExcel(filePath);
    const sheet = workbook.getWorksheet('FillTest');
    
    expect(sheet).toBeDefined();
    if(sheet) {
        // Header Fill
        const headerStyle = getCellStyle(sheet, 1, 1);
        expect(headerStyle.fill).toMatchObject({
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFCCCCCC' }
        });

        // Row Fill (Odd index in data -> Even row in Excel because of header)
        // data[0] (id:1) -> index 0 -> no fill -> Excel Row 2
        // data[1] (id:2) -> index 1 -> fill -> Excel Row 3
        
        const row2Style = getCellStyle(sheet, 2, 1);
        expect(row2Style.fill).toBeUndefined();

        const row3Style = getCellStyle(sheet, 3, 1);
        expect(row3Style.fill).toMatchObject({
             type: 'pattern',
             pattern: 'solid',
             fgColor: { argb: 'FFEFEFEF' }
        });
    }
  });
});
