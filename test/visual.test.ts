import { describe, it } from 'vitest';
import { createWorkbook } from '../src';
import * as path from 'path';
import * as fs from 'fs';

const OUTPUT_DIR = path.join(__dirname, 'output');
if (!fs.existsSync(OUTPUT_DIR)) {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

describe('Visual Verification Gallery', () => {
  
  it('should generate a single gallery workbook with multiple sheets', async () => {
    const filePath = path.join(OUTPUT_DIR, 'visual_gallery.xlsx');
    const workbook = createWorkbook();

    // --- 1. Fonts ---
    workbook.addSheet({
      name: 'Fonts',
      headers: [
        { key: 'label', label: 'Style', width: 15 },
        { 
          key: 'text', 
          label: 'Preview', 
          width: 30,
          style: (val, row) => {
            const s: any = { font: {} };
            if (row.label === 'Bold') s.font.bold = true;
            if (row.label === 'Italic') s.font.italic = true;
            if (row.label === 'Underline') s.font.underline = true;
            if (row.label === 'Strike') s.font.strike = true;
            if (row.label === 'Size 14') s.font.size = 14;
            if (row.label === 'Size 18') s.font.size = 18;
            if (row.label === 'Color Red') s.font.color = '#FF0000';
            if (row.label === 'Color Blue') s.font.color = '#0000FF';
            return s;
          }
        }
      ],
      rows: [
        { label: 'Normal', text: 'The quick brown fox' },
        { label: 'Bold', text: 'The quick brown fox' },
        { label: 'Italic', text: 'The quick brown fox' },
        { label: 'Underline', text: 'The quick brown fox' },
        { label: 'Strike', text: 'The quick brown fox' },
        { label: 'Size 14', text: 'The quick brown fox' },
        { label: 'Size 18', text: 'The quick brown fox' },
        { label: 'Color Red', text: 'The quick brown fox' },
        { label: 'Color Blue', text: 'The quick brown fox' },
      ]
    });

    // --- 2. Fills ---
    workbook.addSheet({
      name: 'Fills',
      headers: [
        { key: 'color', label: 'Color Name', width: 15 },
        { key: 'hex', label: 'Hex Code', width: 15 },
        { 
          key: 'hex', 
          label: 'Preview', 
          width: 15,
          style: (val, row) => ({ fill: { color: row.hex } })
        }
      ],
      rows: [
        { color: 'Red', hex: '#FF0000' },
        { color: 'Green', hex: '#00FF00' },
        { color: 'Blue', hex: '#0000FF' },
        { color: 'Yellow', hex: '#FFFF00' },
        { color: 'Gray', hex: '#CCCCCC' },
      ]
    });

    // --- 3. Alignment ---
    workbook.addSheet({
      name: 'Alignment',
      headers: [
        { key: 'h', label: 'Horizontal', width: 15 },
        { key: 'v', label: 'Vertical', width: 15 },
        { 
          key: 'h', 
          label: 'Preview', 
          width: 20,
          style: (val, row) => ({ 
            alignment: { 
              horizontal: row.h as any, 
              vertical: row.v as any 
            } 
          })
        }
      ],
      rows: [
        { h: 'left', v: 'top' },
        { h: 'center', v: 'top' },
        { h: 'right', v: 'top' },
        { h: 'left', v: 'middle' },
        { h: 'center', v: 'middle' },
        { h: 'right', v: 'middle' },
        { h: 'left', v: 'bottom' },
        { h: 'center', v: 'bottom' },
        { h: 'right', v: 'bottom' },
      ]
    });

    // --- 4. Borders ---
    const borderData = [{ name: 'Item 1' }, { name: 'Item 2' }, { name: 'Item 3' }];
    
    workbook.addSheet({
      name: 'Borders (All)',
      headers: [{ key: 'name', label: 'Name', width: 20 }],
      rows: borderData,
      borders: 'all'
    });

    workbook.addSheet({
      name: 'Borders (Outer)',
      headers: [{ key: 'name', label: 'Name', width: 20 }],
      rows: borderData,
      borders: 'outer'
    });

    workbook.addSheet({
      name: 'Borders (Header)',
      headers: [{ key: 'name', label: 'Name', width: 20 }],
      rows: borderData,
      borders: 'header-body'
    });

    // --- 5. Comprehensive 10x10 ---
    const gridData: any[] = [];
    const categories = ['Electronics', 'Furniture', 'Office', 'Kitchen'];
    const statuses = ['In Stock', 'Low Stock', 'Out of Stock'];

    for (let i = 1; i <= 10; i++) {
      const price = Math.floor(Math.random() * 1000) + 10;
      const qty = Math.floor(Math.random() * 20) + 1;
      gridData.push({
        id: i,
        category: categories[i % categories.length],
        product: `Product ${i} - ${Math.random().toString(36).substring(7)}`,
        date: new Date(2025, 0, i),
        price: price,
        quantity: qty,
        total: price * qty,
        rate: Math.random(),
        status: statuses[i % statuses.length],
        code: `00${i}`.slice(-3)
      });
    }

    workbook.addSheet({
      name: 'Comprehensive',
      title: {
        label: 'Sales Report - January 2025',
        style: { 
          fill: { color: '#4472C4' }, 
          font: { color: '#FFFFFF', bold: true, size: 14 },
          alignment: { horizontal: 'center' }
        }
      },
      headers: [
        { key: 'id', label: 'ID', width: 8 },
        { key: 'category', label: 'Category', width: 15, merge: 'vertical', style: { alignment: { vertical: 'middle' } } },
        { key: 'product', label: 'Product Name', width: 25 },
        { key: 'date', label: 'Date', width: 15, format: 'yyyy-mm-dd' },
        { key: 'price', label: 'Price', width: 12, format: '$#,##0.00' },
        { key: 'quantity', label: 'Qty', width: 8, style: { alignment: { horizontal: 'center' } } },
        { key: 'total', label: 'Total', width: 15, format: '$#,##0.00', style: { font: { bold: true } } },
        { key: 'rate', label: 'Rate', width: 10, format: '0.0%' },
        { 
          key: 'status', 
          label: 'Status', 
          width: 15,
          style: (val) => {
            if (val === 'In Stock') return { font: { color: '#008800' } };
            if (val === 'Low Stock') return { font: { color: '#FFA500' } };
            return { font: { color: '#FF0000' } };
          }
        },
        { key: 'code', label: 'Code', width: 10, style: { alignment: { horizontal: 'center' } } }
      ],
      rows: gridData,
      styles: {
        row: (_, i) => i % 2 === 1 ? { fill: { color: '#F2F2F2' } } : {}
      },
      borders: 'all'
    });

    // Save all in one file
    await workbook.save(filePath);
  });
});
