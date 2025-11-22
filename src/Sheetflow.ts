import ExcelJS from 'exceljs';
import { SheetConfig, HeaderDef, XLStyle, CellValue, StylesConfig } from './types';
import { mapStyle } from './utils/style';

export class XLKit {
  private workbook: ExcelJS.Workbook;

  constructor() {
    this.workbook = new ExcelJS.Workbook();
  }

  addSheet(config: SheetConfig): XLKit {
    // Validate Sheet Name
    if (!config.name) {
      throw new Error('Sheet name is required.');
    }
    if (config.name.length > 31) {
      throw new Error(`Sheet name "${config.name}" exceeds the maximum length of 31 characters.`);
    }
    // Invalid characters: \ / ? * [ ] :
    const invalidChars = /[\\/?*[\]:]/;
    if (invalidChars.test(config.name)) {
      throw new Error(`Sheet name "${config.name}" contains invalid characters (\\ / ? * [ ] :).`);
    }

    const sheet = this.workbook.addWorksheet(config.name);
    const data = config.rows;

    // Handle autoWidth
    const autoWidthConfig = typeof config.autoWidth === 'boolean' 
      ? { enabled: config.autoWidth } 
      : config.autoWidth || {};
    const autoWidthEnabled = autoWidthConfig.enabled !== false;

    // 1. Setup Columns & Headers
    const columns = config.headers.map((header, colIndex) => {
      let width = header.width;

      // Apply autoWidth if no width specified and autoWidth is enabled
      if (!width && autoWidthEnabled) {
        width = 'auto';
      }

      if (width === 'auto') {
        // Get header label text
        const headerText = typeof header.label === 'string' 
          ? header.label 
          : header.label.value;
        
        let maxLen = headerText.length * (autoWidthConfig.headerIncluded !== false ? 1 : 0);
        
        // Check data length
        data.forEach(row => {
          const cellData = row[header.key];
          const val = this.isCellValueWithStyle(cellData) ? cellData.value : cellData;
          const str = val != null ? String(val) : '';
          // Simple full-width check: count as 2 if char code > 255
          let len = 0;
          for (let i = 0; i < str.length; i++) {
            len += str.charCodeAt(i) > 255 ? 2 : 1;
          }
          if (len > maxLen) maxLen = len;
        });

        const padding = autoWidthConfig.padding ?? 2;
        const constant = autoWidthConfig.charWidthConstant ?? 1.2;
        width = (maxLen + padding) * constant;
      }

      // Get header label text for ExcelJS
      const headerText = typeof header.label === 'string' 
        ? header.label 
        : header.label.value;

      return {
        header: headerText,
        key: String(header.key),
        width: typeof width === 'number' ? width : 15
      };
    });
    sheet.columns = columns;

    // 2. Apply Title Rows (if any)
    if (config.title) {
      const titleLabels = Array.isArray(config.title.label) 
        ? config.title.label 
        : [config.title.label];
      
      titleLabels.forEach(titleText => {
        const titleRow = sheet.insertRow(1, [titleText]);
        
        // Merge title across all columns
        if (config.headers.length > 1) {
          sheet.mergeCells(1, 1, 1, config.headers.length);
        }
        
        // Apply title style
        if (config.title!.style) {
          const mappedTitleStyle = mapStyle(config.title!.style);
          titleRow.eachCell((cell) => {
            cell.style = { ...cell.style, ...mappedTitleStyle };
          });
        }
      });
    }

    // Calculate header row index (after title rows)
    const titleRowCount = config.title 
      ? (Array.isArray(config.title.label) ? config.title.label.length : 1) 
      : 0;
    const headerRowIndex = titleRowCount + 1;

    // 3. Apply Header Cell Styles (from headers[].label.style)
    const headerRow = sheet.getRow(headerRowIndex);
    config.headers.forEach((header, colIndex) => {
      if (typeof header.label === 'object' && header.label.style) {
        const cell = headerRow.getCell(colIndex + 1);
        const mappedStyle = mapStyle(header.label.style);
        cell.style = { ...cell.style, ...mappedStyle };
      }
    });

    // 4. Apply Header Row Style (from styles.header)
    if (config.styles?.header) {
      const mappedHeaderStyle = mapStyle(config.styles.header);
      headerRow.eachCell((cell) => {
        cell.style = { ...cell.style, ...mappedHeaderStyle };
      });
    }

    // 5. Apply styles.all to header row
    if (config.styles?.all) {
      const mappedAllStyle = mapStyle(config.styles.all);
      headerRow.eachCell((cell) => {
        cell.style = { ...mappedAllStyle, ...cell.style };
      });
    }

    // 6. Add Data & Apply Styles
    data.forEach((rowData, rowIndex) => {
      const rowValues: any = {};
      
      // Extract values from row data
      config.headers.forEach(header => {
        const cellData = rowData[header.key];
        rowValues[header.key] = this.isCellValueWithStyle(cellData) ? cellData.value : cellData;
      });

      const addedRow = sheet.addRow(rowValues);
      const excelRowIndex = headerRowIndex + rowIndex + 1;

      // Apply styles to each cell
      config.headers.forEach((header, colIndex) => {
        const cell = addedRow.getCell(colIndex + 1);
        const cellData = rowData[header.key];
        const cellValue = this.isCellValueWithStyle(cellData) ? cellData.value : cellData;

        // Apply styles in priority order
        let finalStyle: any = {};

        // 1. styles.all
        if (config.styles?.all) {
          finalStyle = { ...finalStyle, ...mapStyle(config.styles.all) };
        }

        // 2. styles.body
        if (config.styles?.body) {
          finalStyle = { ...finalStyle, ...mapStyle(config.styles.body) };
        }

        // 3. styles.column[key]
        if (config.styles?.column?.[header.key]) {
          finalStyle = { ...finalStyle, ...mapStyle(config.styles.column[header.key]) };
        }

        // 4. styles.row(data, index)
        if (config.styles?.row) {
          const rowStyle = config.styles.row(rowData, rowIndex);
          finalStyle = { ...finalStyle, ...mapStyle(rowStyle) };
        }

        // 5. headers[].style (object or function)
        if (header.style) {
          if (typeof header.style === 'function') {
            const cellStyle = header.style(cellValue, rowData, rowIndex);
            finalStyle = { ...finalStyle, ...mapStyle(cellStyle) };
          } else {
            finalStyle = { ...finalStyle, ...mapStyle(header.style) };
          }
        }

        // 6. rows[].{key}.style (highest priority)
        if (this.isCellValueWithStyle(cellData) && cellData.style) {
          finalStyle = { ...finalStyle, ...mapStyle(cellData.style) };
        }

        cell.style = finalStyle;

        // Apply format
        if (header.format) {
          if (typeof header.format === 'string') {
            cell.numFmt = header.format;
          } else {
            cell.value = header.format(cellValue);
          }
        }
      });
    });

    // 7. Apply Vertical Merges
    config.headers.forEach((header, colIndex) => {
      if (header.merge === 'vertical') {
        let startRow = headerRowIndex + 1; // First data row
        let previousValue: any = null;

        // Iterate from first data row to last
        for (let i = 0; i < data.length; i++) {
          const currentRowIndex = headerRowIndex + i + 1;
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
        const lastRowIndex = headerRowIndex + data.length;
        if (lastRowIndex > startRow) {
          sheet.mergeCells(startRow, colIndex + 1, lastRowIndex, colIndex + 1);
        }
      }
    });

    // 8. Apply Borders
    if (config.borders === 'all') {
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
    } else if (config.borders === 'outer') {
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
    } else if (config.borders === 'header-body') {
      const lastCol = sheet.columnCount;
      for (let c = 1; c <= lastCol; c++) {
        const headerCell = sheet.getCell(headerRowIndex, c);
        headerCell.border = { ...headerCell.border, bottom: { style: 'medium' } };
      }
    }

    return this;
  }

  private isCellValueWithStyle(val: any): val is { value: any; style: XLStyle } {
    return val !== null && 
           typeof val === 'object' && 
           'value' in val && 
           !Array.isArray(val) &&
           !(val instanceof Date);
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
