import { describe, it } from 'vitest';
import { createWorkbook, defineSheet } from '../src';
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
    interface FontData { label: string; text: string }
    const fontData: FontData[] = [
        { label: 'Normal', text: 'The quick brown fox' },
        { label: 'Bold', text: 'The quick brown fox' },
        { label: 'Italic', text: 'The quick brown fox' },
        { label: 'Underline', text: 'The quick brown fox' },
        { label: 'Strike', text: 'The quick brown fox' },
        { label: 'Size 14', text: 'The quick brown fox' },
        { label: 'Size 18', text: 'The quick brown fox' },
        { label: 'Color Red', text: 'The quick brown fox' },
        { label: 'Color Blue', text: 'The quick brown fox' },
    ];
    const fontSheet = defineSheet<FontData>({
      name: 'Fonts',
      columns: [
        { key: 'label', header: 'Style', width: 15 },
        { 
            key: 'text', 
            header: 'Preview', 
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
      ]
    });
    workbook.addSheet(fontSheet, fontData);

    // --- 2. Fills ---
    interface FillData { color: string; hex: string }
    const fillData: FillData[] = [
        { color: 'Red', hex: '#FF0000' },
        { color: 'Green', hex: '#00FF00' },
        { color: 'Blue', hex: '#0000FF' },
        { color: 'Yellow', hex: '#FFFF00' },
        { color: 'Gray', hex: '#CCCCCC' },
    ];
    const fillSheet = defineSheet<FillData>({
      name: 'Fills',
      columns: [
        { key: 'color', header: 'Color Name', width: 15 },
        { key: 'hex', header: 'Hex Code', width: 15 },
        { 
            key: 'hex', 
            header: 'Preview', 
            width: 15,
            style: (val, row) => ({ fill: { color: row.hex } })
        }
      ]
    });
    workbook.addSheet(fillSheet, fillData);

    // --- 3. Alignment ---
    interface AlignData { h: string; v: string }
    const alignData: AlignData[] = [
        { h: 'left', v: 'top' },
        { h: 'center', v: 'top' },
        { h: 'right', v: 'top' },
        { h: 'left', v: 'middle' },
        { h: 'center', v: 'middle' },
        { h: 'right', v: 'middle' },
        { h: 'left', v: 'bottom' },
        { h: 'center', v: 'bottom' },
        { h: 'right', v: 'bottom' },
    ];
    const alignSheet = defineSheet<AlignData>({
      name: 'Alignment',
      columns: [
        { key: 'h', header: 'Horizontal', width: 15 },
        { key: 'v', header: 'Vertical', width: 15 },
        { 
            key: 'h', 
            header: 'Preview', 
            width: 20,
            style: (val, row) => ({ 
                alignment: { 
                    horizontal: row.h as any, 
                    vertical: row.v as any 
                } 
            })
        }
      ],
      rows: { style: () => ({}) }
    });
    workbook.addSheet(alignSheet, alignData);

    // --- 4. Borders ---
    interface BorderData { name: string }
    const borderData: BorderData[] = [{ name: 'Item 1' }, { name: 'Item 2' }, { name: 'Item 3' }];
    
    workbook.addSheet(defineSheet<BorderData>({
        name: 'Borders (All)',
        columns: [{ key: 'name', header: 'Name', width: 20 }],
        borders: 'all'
    }), borderData);

    workbook.addSheet(defineSheet<BorderData>({
        name: 'Borders (Outer)',
        columns: [{ key: 'name', header: 'Name', width: 20 }],
        borders: 'outer'
    }), borderData);

    workbook.addSheet(defineSheet<BorderData>({
        name: 'Borders (Header)',
        columns: [{ key: 'name', header: 'Name', width: 20 }],
        borders: 'header-body'
    }), borderData);

    // --- 5. Comprehensive 10x10 ---
    interface GridData {
        id: number;
        category: string;
        product: string;
        date: Date;
        price: number;
        quantity: number;
        total: number;
        rate: number;
        status: string;
        code: string;
    }
    const gridData: GridData[] = [];
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

    const comprehensiveSheet = defineSheet<GridData>({
        name: 'Comprehensive',
        columns: [
            { key: 'id', header: 'ID', width: 8 },
            { key: 'category', header: 'Category', width: 15, merge: 'vertical', style: { alignment: { vertical: 'middle' } } },
            { key: 'product', header: 'Product Name', width: 25 },
            { key: 'date', header: 'Date', width: 15, format: 'yyyy-mm-dd' },
            { key: 'price', header: 'Price', width: 12, format: '$#,##0.00' },
            { key: 'quantity', header: 'Qty', width: 8, style: { alignment: { horizontal: 'center' } } },
            { key: 'total', header: 'Total', width: 15, format: '$#,##0.00', style: { font: { bold: true } } },
            { key: 'rate', header: 'Rate', width: 10, format: '0.0%' },
            { 
                key: 'status', 
                header: 'Status', 
                width: 15,
                style: (val) => {
                    if (val === 'In Stock') return { font: { color: '#008800' } };
                    if (val === 'Low Stock') return { font: { color: '#FFA500' } };
                    return { font: { color: '#FF0000' } };
                }
            },
            { key: 'code', header: 'Code', width: 10, style: { alignment: { horizontal: 'center' } } }
        ],
        header: {
            style: { 
                fill: { color: '#4472C4' }, 
                font: { color: '#FFFFFF', bold: true, size: 14 },
                alignment: { horizontal: 'center' }
            }
        },
        rows: {
            style: (_, i) => i % 2 === 1 ? { fill: { color: '#F2F2F2' } } : {}
        },
        borders: 'all'
    });
    workbook.addSheet(comprehensiveSheet, gridData);

    // Save all in one file
    await workbook.save(filePath);
  });
});
