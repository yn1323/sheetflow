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

    // 1. Setup Columns
    const columns = def.columns.map((col, colIndex) => {
      let width = col.width;

      // Validate Column Key
      if (col.key === 'style') {
        throw new Error("Column key 'style' is reserved for row styling and cannot be used as a column key.");
      }

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

    // 2. Setup Headers (Multi-line support)
    let headerRowCount = 1;
    if (def.header?.rows) {
      headerRowCount = def.header.rows.length;
      // Clear default header row if we are using custom rows
      sheet.spliceRows(1, 1); 

      def.header.rows.forEach((row, rowIndex) => {
        const currentSheetRowIndex = rowIndex + 1;
        const sheetRow = sheet.getRow(currentSheetRowIndex);
        
        let colIndex = 1;
        row.forEach((cellConfig) => {
           // Find next available cell (skip merged cells)
           while (sheet.getCell(currentSheetRowIndex, colIndex).isMerged) {
             colIndex++;
           }

           const cell = sheet.getCell(currentSheetRowIndex, colIndex);
           
           if (typeof cellConfig === 'string') {
             cell.value = cellConfig;
             colIndex++;
           } else {
             cell.value = cellConfig.value;
             
             if (cellConfig.style) {
                cell.style = { ...cell.style, ...mapStyle(cellConfig.style) };
             }

             const rowSpan = cellConfig.rowSpan || 1;
             const colSpan = cellConfig.colSpan || 1;

             if (rowSpan > 1 || colSpan > 1) {
               sheet.mergeCells(
                 currentSheetRowIndex, 
                 colIndex, 
                 currentSheetRowIndex + rowSpan - 1, 
                 colIndex + colSpan - 1
               );
             }
             colIndex += colSpan;
           }
        });
      });
    }

    // 3. Add Data & Apply Row Styles
    const dataStartRow = headerRowCount + 1;
    
    data.forEach((row, index) => {
      const rowIndex = dataStartRow + index;
      // We can't use addRow easily because it appends to the end, but we might have gaps if we messed with rows?
      // Actually addRow is fine as long as we are consistent.
      // But since we might have complex headers, let's be explicit with getRow to be safe or just use addRow if we know we are at the end.
      // For safety with existing header logic, let's use explicit row rendering for data to ensure alignment.
      
      const sheetRow = sheet.getRow(rowIndex);
      
      def.columns.forEach((col, colIndex) => {
          const cell = sheetRow.getCell(colIndex + 1);
          cell.value = row[col.key] as any;
      });
      
      // Apply row-level style from definition
      if (def.rows?.style) {
        const rowStyle = def.rows.style(row, index);
        const mappedStyle = mapStyle(rowStyle);
        sheetRow.eachCell((cell) => {
          cell.style = { ...cell.style, ...mappedStyle };
        });
      }

      // Apply row-level style from data (if 'style' property exists)
      if ((row as any).style) {
         const dataRowStyle = (row as any).style;
         const mappedStyle = mapStyle(dataRowStyle);
         sheetRow.eachCell((cell) => {
            cell.style = { ...cell.style, ...mappedStyle };
         });
      }

      // Apply column-level conditional styles
      def.columns.forEach((col, colIndex) => {
        if (typeof col.style === 'function') {
          const cell = sheetRow.getCell(colIndex + 1);
          const cellStyle = col.style(row[col.key], row, index);
          cell.style = { ...cell.style, ...mapStyle(cellStyle) };
        }
        
        if (col.format) {
             const cell = sheetRow.getCell(colIndex + 1);
             if (typeof col.format === 'string') {
                 cell.numFmt = col.format;
             } else {
                 cell.value = col.format(row[col.key]);
             }
        }
      });
      
      sheetRow.commit();
    });

    // 4. Apply Header Styles (Global)
    if (def.header?.style) {
      const mappedHeaderStyle = mapStyle(def.header.style);
      // Apply to all header rows
      for (let i = 1; i <= headerRowCount; i++) {
        const row = sheet.getRow(i);
        row.eachCell((cell) => {
           // Merge with existing style (cell specific style takes precedence if we did it right, but here we are applying global header style)
           // Usually global header style is base, cell specific is override. 
           // But here we apply global AFTER. Let's apply it only if no style? 
           // Or just merge.
           cell.style = { ...mappedHeaderStyle, ...cell.style };
        });
      }
    }
    
    // 5. Apply Vertical Merges
    def.columns.forEach((col, colIndex) => {
      if (col.merge === 'vertical') {
        let startRow = dataStartRow; 
        let previousValue: any = null;

        // Iterate from first data row to last
        for (let i = 0; i < data.length; i++) {
          const currentRowIndex = dataStartRow + i;
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
        const lastRowIndex = dataStartRow + data.length; // This is actually the row AFTER the last data row index if we use <
        // Wait, loop goes 0 to length-1.
        // If i=length-1 (last item), we check logic.
        // We need to handle the merge AFTER the loop for the final group.
        if (data.length > 0) {
             const finalRowIndex = dataStartRow + data.length - 1;
             if (finalRowIndex > startRow) {
                 sheet.mergeCells(startRow, colIndex + 1, finalRowIndex, colIndex + 1);
             }
        }
      }
    });

    // 6. Apply Horizontal Merges
    // We iterate row by row
    for (let i = 0; i < data.length; i++) {
        const currentRowIndex = dataStartRow + i;
        const row = sheet.getRow(currentRowIndex);
        
        let startCol = 1;
        let previousValue: any = null;
        let merging = false;

        // We need to check which columns are candidates for horizontal merge.
        // A simple approach: iterate columns. If current col has merge='horizontal', try to merge with next.
        // But we need to group them.
        
        for (let c = 0; c < def.columns.length; c++) {
            const colDef = def.columns[c];
            const currentCell = row.getCell(c + 1);
            const currentValue = currentCell.value;

            if (colDef.merge === 'horizontal') {
                if (!merging) {
                    merging = true;
                    startCol = c + 1;
                    previousValue = currentValue;
                } else {
                    if (currentValue !== previousValue) {
                        // End of a merge group
                        if ((c + 1) - 1 > startCol) {
                            sheet.mergeCells(currentRowIndex, startCol, currentRowIndex, c);
                        }
                        // Start new group
                        startCol = c + 1;
                        previousValue = currentValue;
                    }
                }
            } else {
                if (merging) {
                    // End of merge group because this column is not mergeable
                    if ((c + 1) - 1 > startCol) {
                        sheet.mergeCells(currentRowIndex, startCol, currentRowIndex, c);
                    }
                    merging = false;
                }
            }
        }
        // Check at end of row
        if (merging) {
             const lastCol = def.columns.length;
             if (lastCol > startCol) {
                 sheet.mergeCells(currentRowIndex, startCol, currentRowIndex, lastCol);
             }
        }
    }

    // 7. Apply Borders
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
         // Apply to the last row of the header
         for (let c = 1; c <= lastCol; c++) {
             const headerCell = sheet.getCell(headerRowCount, c);
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
