import { describe, it, expect } from 'vitest';
import { createWorkbook, defineSheet } from '../src';

describe('API Validation', () => {
  it('should throw error if column key is "style"', () => {
    expect(() => {
      const sheet = defineSheet({
        name: 'Invalid',
        columns: [
          { key: 'style', header: 'Style', width: 10 }
        ]
      });
      const workbook = createWorkbook();
      workbook.addSheet(sheet, []);
    }).toThrow("Column key 'style' is reserved for row styling and cannot be used as a column key.");
  });

  it('should allow "style" property in row data for styling', async () => {
    const workbook = createWorkbook();
    const sheet = defineSheet<{ name: string; style?: any }>({
      name: 'Row Style',
      columns: [
        { key: 'name', header: 'Name', width: 20 }
      ]
    });

    const data = [
      { name: 'Normal' },
      { name: 'Red', style: { font: { color: '#FF0000' } } }
    ];

    workbook.addSheet(sheet, data);
    // No error should be thrown
  });
});
