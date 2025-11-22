# xlkit

<p align="center">
  <img src="./logo.png" alt="xlkit Logo" width="200" />
</p>

A declarative, schema-based wrapper for [ExcelJS](https://github.com/exceljs/exceljs).  
Define your Excel structure with a simple schema and let xlkit handle the styling, formatting, and layout.

## Features

- 📝 **Declarative Schema**: Define data and schema in one place.
- 🎨 **Flexible Styling**: Apply styles at 7 different levels (title, header, row, column, cell).
- 🔗 **Auto Merge**: Automatically merge vertical cells with the same value.
- 📏 **Auto Width**: Smart column width calculation based on content (including full-width chars).
- 🌈 **Hex Colors**: Use standard 6-digit hex codes (`#FF0000`) directly.
- 🌐 **Universal**: Works in Node.js (file output) and Browser/Frontend (`Uint8Array` output).

## Installation

```bash
npm install xlkit
```

## Quick Start

```typescript
import { createWorkbook } from 'xlkit';

await createWorkbook().addSheet({
  name: 'Users',
  headers: [
    { key: 'id', label: 'ID', width: 10 },
    { key: 'name', label: 'Name', width: 20 },
    { 
      key: 'role', 
      label: 'Role', 
      width: 'auto', 
      merge: 'vertical' 
    },
    { 
      key: 'isActive', 
      label: 'Status', 
      format: (val) => val ? 'Active' : 'Inactive',
      style: (val) => ({ font: { color: val ? '#00AA00' : '#FF0000' } })
    }
  ],
  rows: [
    { id: 1, name: 'Alice', role: 'Admin', isActive: true },
    { id: 2, name: 'Bob', role: 'User', isActive: true },
    { id: 3, name: 'Charlie', role: 'User', isActive: false }
  ],
  borders: 'outer'
}).save('users.xlsx');
```

## API Reference

### 1. Basic Structure

```typescript
createWorkbook().addSheet({
  name: string,              // Sheet name (required)
  headers: HeaderDef[],      // Column definitions (required)
  rows: any[],               // Data rows (required)
  title?: TitleConfig,       // Title row (optional)
  styles?: StylesConfig,     // Global styles (optional)
  borders?: 'all' | 'outer' | 'header-body' | 'none',
  autoWidth?: boolean | { ... }
})
```

### 2. Headers (`headers`)

Define columns with the `headers` array.

```typescript
headers: [
  { 
    key: 'age',                    // Data property key
    label: 'Age',                  // Header text
    width: 10,                     // Column width (number or 'auto')
    merge: 'vertical',             // Auto-merge vertically
    format: '$#,##0',              // Number/date format
    style: { ... }                 // Fixed column style
  },
  {
    key: 'salary',
    label: 'Salary',
    style: (val, row, index) => { // Conditional style (function)
      return val > 100000 ? { font: { color: '#FF0000' } } : {};
    }
  }
]
```

**Header Cell Styling:**
```typescript
headers: [
  { 
    key: 'age', 
    label: { value: 'Age', style: { font: { bold: true } } }  // Style header cell
  }
]
```

### 3. Data Rows (`rows`)

Define data and cell-level styles.

```typescript
rows: [
  { age: 18, name: "Mary" },  // Simple values
  { 
    age: 25, 
    name: { value: "Tom", style: { font: { bold: true } } }  // Cell with style
  }
]
```

### 4. Title Row (`title`)

Add a title row at the top of the sheet.

```typescript
title: {
  label: 'Employee List 2025',  // Or array: ['Title 1', 'Title 2']
  style: { 
    fill: { color: '#4472C4' }, 
    font: { color: '#FFFFFF', bold: true, size: 14 },
    alignment: { horizontal: 'center' }
  }
}
```

### 5. Global Styles (`styles`)

Apply styles at 7 different priority levels.

```typescript
styles: {
  all: { font: { name: 'Arial', size: 11 } },  // Default for all
  header: { fill: { color: '#EEEEEE' }, font: { bold: true } },  // Header row
  body: { alignment: { vertical: 'middle' } },  // Body area
  row: (data, index) => {  // Row-level (dynamic)
    return index % 2 === 1 ? { fill: { color: '#F2F2F2' } } : {};
  },
  column: {  // Column-level
    age: { alignment: { horizontal: 'center' } },
    name: { font: { bold: true } }
  }
}
```

**Style Priority (Header Row):**
1. `styles.all` → 2. `styles.header` → 3. `headers[].label.style`

**Style Priority (Data Rows):**
1. `styles.all` → 2. `styles.body` → 3. `styles.column[key]` → 4. `styles.row()` → 5. `headers[].style` → 6. `rows[].{key}.style`

### 6. Borders (`borders`)

Apply border presets to the entire sheet.

- **`'all'`**: Grid borders on all cells
- **`'outer'`**: Border only on outer edges
- **`'header-body'`**: Thick line below header
- **`'none'`**: No borders (default)

```typescript
{
  borders: 'all'
}
```

### 7. Auto Width (`autoWidth`)

```typescript
// Method 1: Auto-adjust all columns
{ autoWidth: true }

// Method 2: Detailed configuration
{ 
  autoWidth: {
    enabled: true,
    padding: 2,
    headerIncluded: true,
    charWidthConstant: 1.2
  }
}

// Method 3: Individual width takes priority
{
  headers: [
    { key: 'age', label: 'Age', width: 10 },  // Fixed width
    { key: 'name', label: 'Name' }  // Auto-adjust
  ],
  autoWidth: true
}
```

### 8. Browser Download

```typescript
// Node.js
await createWorkbook().addSheet({ ... }).save('output.xlsx');

// Browser
await createWorkbook().addSheet({ ... }).download('output.xlsx');
```

### 9. Timeout Configuration

Default 10-second timeout to prevent freezing with large datasets.

```typescript
// Default (10 seconds)
await createWorkbook().addSheet({ ... }).save('output.xlsx');

// Custom timeout (30 seconds)
await createWorkbook().addSheet({ ... }).save('output.xlsx', { timeout: 30000 });
```

> **Recommendation**: Default setting works well for datasets under 100,000 rows.

## Complete Example

```typescript
await createWorkbook().addSheet({
  name: 'Employees',
  title: {
    label: 'Employee List 2025',
    style: { 
      fill: { color: '#4472C4' }, 
      font: { color: '#FFFFFF', bold: true, size: 14 },
      alignment: { horizontal: 'center' }
    }
  },
  headers: [
    { 
      key: 'dept', 
      label: 'Department', 
      merge: 'vertical',
      style: { alignment: { vertical: 'middle', horizontal: 'center' } }
    },
    { key: 'name', label: 'Name', width: 20 },
    { 
      key: 'salary', 
      label: 'Salary',
      format: '$#,##0',
      style: (val) => val > 100000 ? { font: { color: '#FF0000', bold: true } } : {}
    }
  ],
  rows: [
    { dept: 'Engineering', name: 'Alice', salary: 120000 },
    { dept: 'Engineering', name: 'Bob', salary: 80000 },
    { dept: 'Sales', name: { value: 'Charlie', style: { font: { bold: true } } }, salary: 95000 }
  ],
  styles: {
    row: (_, index) => index % 2 === 1 ? { fill: { color: '#F2F2F2' } } : {}
  },
  borders: 'all',
  autoWidth: true
}).save('employees.xlsx');
```

## License

MIT
