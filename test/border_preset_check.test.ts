import { describe, it, expect } from 'vitest';
import { createWorkbook, defineSheet } from '../src';
import { readExcel } from './utils';
import * as path from 'path';
import * as fs from 'fs';

const OUTPUT_DIR = path.join(__dirname, 'output');
if (!fs.existsSync(OUTPUT_DIR)) {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

describe('Border Preset in Cell Style', () => {
  it('should check if border presets work in cell style', async () => {
    const filePath = path.join(OUTPUT_DIR, 'border_preset_cell.xlsx');
    interface Data { val: string }
    const data: Data[] = [{ val: 'A' }];

    const sheetDef = defineSheet<Data>({
      name: 'BorderPreset',
      columns: [
        { 
            key: 'val', 
            header: 'Value',
            // Trying to use a preset string in cell style
            style: { border: 'all' as any } 
        }
      ]
    });

    await createWorkbook().addSheet(sheetDef, data).save(filePath);

    const workbook = await readExcel(filePath);
    const sheet = workbook.getWorksheet('BorderPreset');
    
    if(sheet) {
        const cell = sheet.getCell(2, 1);
        // If 'all' worked, it should have borders. If not, it might be undefined or empty.
        // Based on code analysis, mapStyle ignores string borders, so we expect this to FAIL if we assert it works.
        // We want to know the actual behavior.
        // Based on code analysis, mapStyle ignores string borders, so we expect this to be undefined.
        expect(cell.border).toBeUndefined();
    }
  });
});
