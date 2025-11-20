import ExcelJS from 'exceljs';
import { SheetDef, ColumnDef, XLStyle } from './types';
import { mapStyle } from './utils/style';

export class XLKit {
  private workbook: ExcelJS.Workbook;

  constructor() {
    this.workbook = new ExcelJS.Workbook();
  }

  addSheet<T>(def: SheetDef<T>, data: T[]): XLKit {
    // Validate Sheet Name
    if (!def.name) {
        throw new Error('Sheet name is required.');
    }
    if (def.name.length > 31) {
        throw new Error(`Sheet name "${def.name}" exceeds the maximum length of 31 characters.`);
    }
    // Invalid characters: \ / ? * [ ] :
    const invalidChars = /[\\/?*[\]:]/;
    if (invalidChars.test(def.name)) {
        throw new Error(`Sheet name "${def.name}" contains invalid characters (\\ / ? * [ ] :).`);
    }

    const sheet = this.workbook.addWorksheet(def.name);

    // 1. Setup Columns & Headers
    const columns = def.columns.map((col, colIndex) => {
      let width = col.width;

      if (width === 'auto') {
        let maxLen = col.header.length * (def.autoWidth?.headerIncluded !== false ? 1 : 0);
        
        // Check data length (sample first 100 rows for performance if needed, currently all)
        data.forEach(row => {
          const val = row[col.key];
          const str = val != null ? String(val) : '';
          // Simple full-width check: count as 2 if char code > 255
          let len = 0;
          for (let i = 0; i < str.length; i++) {
            len += str.charCodeAt(i) > 255 ? 2 : 1;
          }
          if (len > maxLen) maxLen = len;
        });

        const padding = def.autoWidth?.padding ?? 2;
        const constant = def.autoWidth?.charWidthConstant ?? 1.2;
        width = (maxLen + padding) * constant;
      }

      return {
        header: col.header,
        key: String(col.key),
        width: typeof width === 'number' ? width : 15, 
        style: col.style && typeof col.style === 'object' ? mapStyle(col.style) : undefined
      };
    });
    sheet.columns = columns;

    // 2. Add Data & Apply Row Styles
    data.forEach((row, rowIndex) => {
      const addedRow = sheet.addRow(row);
      
      // Apply row-level style
      if (def.rows?.style) {
        const rowStyle = def.rows.style(row, rowIndex);
        const mappedStyle = mapStyle(rowStyle);
        addedRow.eachCell((cell) => {
          cell.style = { ...cell.style, ...mappedStyle };
        });
      }

      // Apply column-level conditional styles
      def.columns.forEach((col, colIndex) => {
        if (typeof col.style === 'function') {
          const cell = addedRow.getCell(colIndex + 1);
          const cellStyle = col.style(row[col.key], row, rowIndex);
          cell.style = { ...cell.style, ...mapStyle(cellStyle) };
        }
        
        if (col.format) {
             const cell = addedRow.getCell(colIndex + 1);
             if (typeof col.format === 'string') {
                 cell.numFmt = col.format;
             } else {
                 cell.value = col.format(row[col.key]);
             }
        }
      });
    });

    // 3. Apply Header Styles
    if (def.header?.style) {
      const headerRow = sheet.getRow(1);
      const mappedHeaderStyle = mapStyle(def.header.style);
      headerRow.eachCell((cell) => {
        cell.style = { ...cell.style, ...mappedHeaderStyle };
      });
    }
    
    // 4. Apply Vertical Merges
    def.columns.forEach((col, colIndex) => {
      if (col.merge === 'vertical') {
        let startRow = 2; // 1-based, skip header (row 1)
        let previousValue: any = null;

        // Iterate from first data row to last
        for (let i = 0; i < data.length; i++) {
          const currentRowIndex = i + 2;
          const cell = sheet.getCell(currentRowIndex, colIndex + 1);
          const currentValue = cell.value;

          if (i === 0) {
            previousValue = currentValue;
            continue;
          }

          // If value changed or it's the last row, process the merge
          if (currentValue !== previousValue) {
            if (currentRowIndex - 1 > startRow) {
              sheet.mergeCells(startRow, colIndex + 1, currentRowIndex - 1, colIndex + 1);
            }
            startRow = currentRowIndex;
            previousValue = currentValue;
          }
        }
        
        // Handle the last group
        const lastRowIndex = data.length + 1;
        if (lastRowIndex > startRow) {
             sheet.mergeCells(startRow, colIndex + 1, lastRowIndex, colIndex + 1);
        }
      }
    });

    // 5. Apply Borders
    if (def.borders === 'all') {
        sheet.eachRow((row) => {
            row.eachCell((cell) => {
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            });
        });
    } else if (def.borders === 'outer') {
        const lastRow = sheet.rowCount;
        const lastCol = sheet.columnCount;
        
        // Top & Bottom
        for (let c = 1; c <= lastCol; c++) {
            const topCell = sheet.getCell(1, c);
            topCell.border = { ...topCell.border, top: { style: 'thin' } };
            const bottomCell = sheet.getCell(lastRow, c);
            bottomCell.border = { ...bottomCell.border, bottom: { style: 'thin' } };
        }
        // Left & Right
        for (let r = 1; r <= lastRow; r++) {
            const leftCell = sheet.getCell(r, 1);
            leftCell.border = { ...leftCell.border, left: { style: 'thin' } };
            const rightCell = sheet.getCell(r, lastCol);
            rightCell.border = { ...rightCell.border, right: { style: 'thin' } };
        }
    } else if (def.borders === 'header-body') {
         const lastCol = sheet.columnCount;
         for (let c = 1; c <= lastCol; c++) {
             const headerCell = sheet.getCell(1, c);
             headerCell.border = { ...headerCell.border, bottom: { style: 'medium' } };
         }
    }

    return this;
  }

  async save(path: string, options?: { timeout?: number }): Promise<void> {
    if (!path || path.trim() === '') {
        throw new Error('File path cannot be empty.');
    }
    if (typeof process !== 'undefined' && process.versions && process.versions.node) {
      const timeout = options?.timeout ?? 10000; // Default 10s
      
      const writePromise = this.workbook.xlsx.writeFile(path);
      
      const timeoutPromise = new Promise<void>((_, reject) => {
          setTimeout(() => reject(new Error(`Operation timed out after ${timeout}ms`)), timeout);
      });

      await Promise.race([writePromise, timeoutPromise]);
    } else {
      throw new Error('File system access is only available in Node.js environment. Use saveToBuffer() instead.');
    }
  }
  
  async saveToBuffer(options?: { timeout?: number }): Promise<Uint8Array> {
      const timeout = options?.timeout ?? 10000; // Default 10s

      const writePromise = this.workbook.xlsx.writeBuffer();
      
      const timeoutPromise = new Promise<ExcelJS.Buffer>((_, reject) => {
          setTimeout(() => reject(new Error(`Operation timed out after ${timeout}ms`)), timeout);
      });

      const buffer = await Promise.race([writePromise, timeoutPromise]);
      return new Uint8Array(buffer as ArrayBuffer);
  }

  async download(filename: string, options?: { timeout?: number }): Promise<void> {
    if (typeof window === 'undefined' || typeof document === 'undefined') {
      throw new Error('download() is only available in browser environment. Use save() for Node.js or saveToBuffer() for custom handling.');
    }

    const buffer = await this.saveToBuffer(options);
    const blob = new Blob([buffer.buffer as ArrayBuffer], { 
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
  }

}

export function createWorkbook(): XLKit {
  return new XLKit();
}

export function defineSheet<T>(def: SheetDef<T>): SheetDef<T> {
  return def;
}
