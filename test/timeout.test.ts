import { describe, it, expect } from 'vitest';
import { createWorkbook } from '../src';
import * as path from 'path';
import * as fs from 'fs';

const OUTPUT_DIR = path.join(__dirname, 'output');
if (!fs.existsSync(OUTPUT_DIR)) {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

describe('Timeout Functionality', () => {
  it('should timeout if operation takes too long', async () => {
    const filePath = path.join(OUTPUT_DIR, 'timeout.xlsx');
    
    // Create enough data to take some time
    const data = Array.from({ length: 1000 }, (_, i) => ({ id: i, name: `Item ${i}` }));

    const sf = createWorkbook().addSheet({
      name: 'Timeout',
      headers: [
        { key: 'id', label: 'ID' },
        { key: 'name', label: 'Name' }
      ],
      rows: data
    });

    // Set timeout to 1ms to ensure it fails
    await expect(sf.save(filePath, { timeout: 1 })).rejects.toThrow(/Operation timed out/);
  });

  it('should respect custom timeout', async () => {
    const filePath = path.join(OUTPUT_DIR, 'custom_timeout.xlsx');

    // Should pass with sufficient timeout
    await expect(createWorkbook().addSheet({
      name: 'Timeout',
      headers: [{ key: 'id', label: 'ID' }, { key: 'name', label: 'Name' }],
      rows: [{ id: 1, name: 'Test' }]
    }).save(filePath, { timeout: 5000 })).resolves.not.toThrow();
  });
});
