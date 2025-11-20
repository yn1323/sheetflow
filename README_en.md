# xlkit

<p align="center">
  <img src="./logo.png" alt="xlkit Logo" width="200" />
</p>

A declarative, schema-based wrapper for [ExcelJS](https://github.com/exceljs/exceljs).  
Define your Excel structure with a simple schema and let xlkit handle the styling, formatting, and layout.

## Features

- 📝 **Declarative Schema**: Define columns, headers, and styles in one place.
- 🎨 **Easy Styling**: Apply styles to headers, rows, columns, or specific cells conditionally.
- 🔗 **Auto Merge**: Automatically merge vertical cells with the same value.
- 📏 **Auto Width**: Smart column width calculation based on content (including full-width chars).
- 🌈 **Hex Colors**: Use standard 6-digit hex codes (`#FF0000`) directly.
- 🌐 **Universal**: Works in Node.js (file output) and Browser/Frontend (`Uint8Array` output).

## Installation

```bash
npm install xlkit
```

## Usage

```typescript
import { createWorkbook, defineSheet } from 'xlkit';

// 1. Define your data type
interface User {
  id: number;
  name: string;
  role: string;
  isActive: boolean;
}

// 2. Define the sheet schema
const userSheet = defineSheet<User>({
  name: 'Users',
  columns: [
    { key: 'id', header: 'ID', width: 10 },
    { 
      key: 'name', 
      header: 'Name', 
      width: 20, 
      style: { font: { bold: true } } 
    },
    { 
      key: 'role', 
      header: 'Role', 
      width: 'auto',
      merge: 'vertical', // Auto-merge same values vertically
      style: { alignment: { vertical: 'middle', horizontal: 'center' } }
    },
    { 
      key: 'isActive', 
      header: 'Status', 
      // Conditional styling
      style: (val) => ({ 
        font: { color: val ? '#00AA00' : '#FF0000' } 
      }),
      format: (val) => val ? 'Active' : 'Inactive'
    }
  ],
  header: {
    rows: ['User List'],
    style: { 
      fill: { color: '#EEEEEE' },
      font: { bold: true }
    },
    borders: 'header-body'
  },
  borders: 'outer'
});

// 3. Generate Excel
const users = [
  { id: 1, name: 'Alice', role: 'Admin', isActive: true },
  { id: 2, name: 'Bob', role: 'User', isActive: true },
  { id: 3, name: 'Charlie', role: 'User', isActive: false },
];

// Node.js
await createWorkbook()
  .addSheet(userSheet, users)
  .save('users.xlsx');

// Frontend / Universal
const buffer = await createWorkbook()
  .addSheet(userSheet, users)
  .saveToBuffer();
// -> Returns Uint8Array
```

### Browser Download

In browser environments, you can easily download Excel files using the `download()` method.

```typescript
// Node.js
await createWorkbook().addSheet(sheet, data).save('output.xlsx');

// Browser
await createWorkbook().addSheet(sheet, data).download('output.xlsx');
```

Internally, it calls `saveToBuffer()`, creates a Blob, and automatically triggers the download.

### Timeout Configuration

To prevent freezing with large datasets, `save()`, `saveToBuffer()`, and `download()` have a default 10-second timeout.

```typescript
// Default (10 seconds)
await createWorkbook().addSheet(sheet, data).save('output.xlsx');

// Custom timeout (30 seconds)
await createWorkbook().addSheet(sheet, data).save('output.xlsx', { timeout: 30000 });

// Browser environment
await createWorkbook().addSheet(sheet, data).download('output.xlsx', { timeout: 15000 });
```

> **Recommendation**: The default setting works well for datasets under 100,000 rows. For larger datasets, consider splitting files or using streaming approaches.

## License

MIT
