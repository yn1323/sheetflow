# xlkit

<p align="center">
  <img src="./logo.png" alt="xlkit Logo" width="200" />
</p>

[ExcelJS](https://github.com/exceljs/exceljs) のための宣言的スキーマベースラッパーです。
シンプルなスキーマでExcelの構造を定義するだけで、スタイル、フォーマット、レイアウトをxlkitが自動で処理します。

[English README](./README_en.md)

## 特徴

- 📝 **宣言的スキーマ**: データとスキーマを一箇所で定義。
- 🎨 **柔軟なスタイル設定**: タイトル、ヘッダー、行、列、セルの7段階でスタイルを適用可能。
- 🔗 **自動結合**: 同じ値を持つ縦方向のセルを自動的に結合 (`merge: 'vertical'`)。
- 📏 **自動列幅**: コンテンツ（全角文字を含む）に基づいて列幅をスマートに計算。
- 🌈 **Hexカラー**: 標準的な6桁のHexコード（`#FF0000`）を直接使用可能。
- 🌐 **ユニバーサル**: Node.js（ファイル出力）とブラウザ/フロントエンド（`Uint8Array` 出力）の両方で動作。

## インストール

```bash
npm install xlkit
```

## クイックスタート

```typescript
import { createWorkbook } from 'xlkit';

await createWorkbook().addSheet({
  name: 'Users',
  headers: [
    { key: 'id', label: 'ID', width: 10 },
    { key: 'name', label: '氏名', width: 20 },
    { 
      key: 'role', 
      label: '役割', 
      width: 'auto', 
      merge: 'vertical' 
    },
    { 
      key: 'isActive', 
      label: 'ステータス', 
      format: (val) => val ? '有効' : '無効',
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

## 詳細リファレンス

### 1. 基本構造

```typescript
createWorkbook().addSheet({
  name: string,              // シート名（必須）
  headers: HeaderDef[],      // 列定義（必須）
  rows: any[],               // データ行（必須）
  title?: TitleConfig,       // タイトル行（オプション）
  styles?: StylesConfig,     // 全体スタイル設定（オプション）
  borders?: 'all' | 'outer' | 'header-body' | 'none',
  autoWidth?: boolean | { ... }
})
```

### 2. ヘッダー定義 (`headers`)

列の定義は `headers` 配列で行います。

```typescript
headers: [
  { 
    key: 'age',                    // データのプロパティキー
    label: '年齢',                 // ヘッダーテキスト
    width: 10,                     // 列幅（数値または'auto'）
    merge: 'vertical',             // 縦方向の自動結合
    format: '$#,##0',              // 数値/日付フォーマット
    style: { ... }                 // 列全体のスタイル（固定）
  },
  {
    key: 'salary',
    label: '給与',
    style: (val, row, index) => { // 条件付きスタイル（関数）
      return val > 100000 ? { font: { color: '#FF0000' } } : {};
    }
  }
]
```

**ヘッダーセルのスタイル指定:**
```typescript
headers: [
  { 
    key: 'age', 
    label: { value: '年齢', style: { font: { bold: true } } }  // ヘッダーセルにスタイル
  }
]
```

### 3. データ行 (`rows`)

データとセルレベルのスタイルを定義します。

```typescript
rows: [
  { age: 18, name: "Mary" },  // シンプルな値
  { 
    age: 25, 
    name: { value: "Tom", style: { font: { bold: true } } }  // セルにスタイル
  }
]
```

### 4. タイトル行 (`title`)

シートの最上部にタイトル行を追加できます。

```typescript
title: {
  label: '従業員リスト 2025',  // または配列: ['タイトル1', 'タイトル2']
  style: { 
    fill: { color: '#4472C4' }, 
    font: { color: '#FFFFFF', bold: true, size: 14 },
    alignment: { horizontal: 'center' }
  }
}
```

### 5. 全体スタイル設定 (`styles`)

7段階の優先順位でスタイルを適用できます。

```typescript
styles: {
  all: { font: { name: 'Arial', size: 11 } },  // 全体のデフォルト
  header: { fill: { color: '#EEEEEE' }, font: { bold: true } },  // ヘッダー行全体
  body: { alignment: { vertical: 'middle' } },  // ボディ全体
  row: (data, index) => {  // 行全体（動的）
    return index % 2 === 1 ? { fill: { color: '#F2F2F2' } } : {};
  },
  column: {  // 列全体
    age: { alignment: { horizontal: 'center' } },
    name: { font: { bold: true } }
  }
}
```

**スタイル適用の優先順位（ヘッダー行）:**
1. `styles.all` → 2. `styles.header` → 3. `headers[].label.style`

**スタイル適用の優先順位（データ行）:**
1. `styles.all` → 2. `styles.body` → 3. `styles.column[key]` → 4. `styles.row()` → 5. `headers[].style` → 6. `rows[].{key}.style`

### 6. 罫線 (`borders`)

シート全体の罫線プリセットを指定できます。

- **`'all'`**: すべてのセルに格子状の罫線
- **`'outer'`**: データ領域の外枠のみ
- **`'header-body'`**: ヘッダー行の下に太めの線
- **`'none'`**: 罫線なし（デフォルト）

```typescript
{
  borders: 'all'
}
```

### 7. 列幅自動調整 (`autoWidth`)

```typescript
// 方法1: 全列を自動調整
{ autoWidth: true }

// 方法2: 詳細設定
{ 
  autoWidth: {
    enabled: true,
    padding: 2,
    headerIncluded: true,
    charWidthConstant: 1.2
  }
}

// 方法3: 個別指定が優先
{
  headers: [
    { key: 'age', label: '年齢', width: 10 },  // 固定幅
    { key: 'name', label: '名前' }  // 自動調整
  ],
  autoWidth: true
}
```

### 8. ブラウザ環境でのダウンロード

```typescript
// Node.js環境
await createWorkbook().addSheet({ ... }).save('output.xlsx');

// ブラウザ環境
await createWorkbook().addSheet({ ... }).download('output.xlsx');
```

### 9. タイムアウト設定

大量データ処理時のフリーズを防ぐため、デフォルトで10秒のタイムアウトが設定されています。

```typescript
// デフォルト（10秒）
await createWorkbook().addSheet({ ... }).save('output.xlsx');

// カスタムタイムアウト（30秒）
await createWorkbook().addSheet({ ... }).save('output.xlsx', { timeout: 30000 });
```

> **推奨**: 10万行以下のデータであればデフォルト設定で問題ありません。

## 完全な例

```typescript
await createWorkbook().addSheet({
  name: 'Employees',
  title: {
    label: '従業員リスト 2025',
    style: { 
      fill: { color: '#4472C4' }, 
      font: { color: '#FFFFFF', bold: true, size: 14 },
      alignment: { horizontal: 'center' }
    }
  },
  headers: [
    { 
      key: 'dept', 
      label: '部署', 
      merge: 'vertical',
      style: { alignment: { vertical: 'middle', horizontal: 'center' } }
    },
    { key: 'name', label: '名前', width: 20 },
    { 
      key: 'salary', 
      label: '給与',
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

## ライセンス

MIT
