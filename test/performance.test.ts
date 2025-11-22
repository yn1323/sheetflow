import { describe, it, expect } from 'vitest';
import { createWorkbook } from '../src';
import * as path from 'path';
import * as fs from 'fs';

const OUTPUT_DIR = path.join(__dirname, 'output');
if (!fs.existsSync(OUTPUT_DIR)) {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

describe('Performance', () => {
  it('should handle 50,000 rows with 20 columns (Wide Dataset)', async () => {
    const filePath = path.join(OUTPUT_DIR, 'large_wide_50k.xlsx');
    
    // Generate 20 columns
    const headers = Array.from({ length: 20 }, (_, i) => ({
      key: `col_${i}`,
      label: `Column ${i}`,
      width: 15
    }));

    // Generate data
    const data: any[] = [];
    for (let i = 0; i < 50000; i++) {
      const row: any = {};
      for (let j = 0; j < 20; j++) {
        row[`col_${j}`] = `Val ${i}-${j}`;
      }
      data.push(row);
    }

    console.log('Starting 50k row * 20 column generation (1M cells)...');
    const start = Date.now();
    await createWorkbook()
      .addSheet({
        name: 'WideData',
        headers: headers,
        rows: data
      })
      .save(filePath);
    const end = Date.now();
    
    console.log(`Generated 50,000 rows (20 cols) in ${end - start}ms`);
    expect(end - start).toBeLessThan(10000); 
  }, 30000);
});
