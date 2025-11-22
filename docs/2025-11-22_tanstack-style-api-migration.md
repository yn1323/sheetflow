# TanStack TableスタイルへのAPI移行

現在の `defineSheet` APIから、TanStack Tableにインスパイアされた `headers` と `rows` プロパティを持つAPIへの完全移行。

## ユーザーレビュー必須

> [!WARNING]
> **破壊的変更**: これはxlkitを使用する既存のコードをすべて壊す完全なAPIの刷新です。
> 
> - 現在のバージョン: `1.0.3`
> - この変更後は `1.1.0` (マイナーバージョン) にバージョンアップ
> - すべての既存ユーザーがコードを移行する必要があります

## 新しいAPI設計の全体像

```typescript
createWorkbook().addSheet({
  name: 'Employees',
  
  // タイトル設定
  title: {
    label: '従業員リスト 2025',
    style: { fill: { color: '#4472C4' }, font: { bold: true, color: '#FFFFFF' } }
  },
  
  // ヘッダー定義（列定義）
  headers: [
    { 
      key: 'age', 
      label: { value: '年齢', style: { font: { bold: true, color: '#0000FF' } } },  // ヘッダーセルにスタイル
      width: 10,
      // 方法1: 固定スタイル（オブジェクト）
      style: { alignment: { horizontal: 'center' } }
    },
    { 
      key: 'salary',
      label: '給与',  // シンプルな文字列
      // 方法2: 条件付きスタイル（関数）
      style: (val, row, index) => val > 100000 ? { font: { color: '#FF0000' } } : {}
    },
    { key: 'name', label: '名前' },
    { key: 'dept', label: '部署', merge: 'vertical' }
  ],
  
  // データ行（セルレベルのスタイルも含む）
  rows: [
    { 
      age: 18, 
      name: "Mary",  // シンプルな値
      dept: { value: "Engineering", style: { font: { bold: true } } }  // セルスタイル付き
    },
    { 
      age: 25, 
      name: { value: "Tom", style: { fill: { color: '#FFFF00' } } },  // セルスタイル付き
      dept: "Engineering"
    }
  ],
  
  // 全体スタイル設定
  styles: {
    all: { font: { name: 'Arial', size: 11 } },  // 全体のデフォルト
    header: { fill: { color: '#EEEEEE' }, font: { bold: true } },  // ヘッダー行全体
    body: { alignment: { vertical: 'middle' } },  // ボディ全体
    row: (data, index) => index % 2 === 1 ? { fill: { color: '#F2F2F2' } } : {},  // 行全体
    column: {  // 列全体
      age: { alignment: { horizontal: 'center' } },
      name: { font: { bold: true } }
    }
  },
  
  // 罫線設定
  borders: 'all',  // または 'outer', 'header-body', 'none'
  
  // 列幅自動調整
  autoWidth: true
  
}).save('output.xlsx');
```

## スタイル付与の優先順位

スタイルは以下の優先順位で適用されます（下に行くほど優先度が高い）：

**ヘッダー行の場合:**
1. `styles.all` - 全体のデフォルト
2. `styles.header` - ヘッダー行全体
3. `headers[].label.style` - ヘッダーセル単位のスタイル（最優先）

**データ行の場合:**
1. `styles.all` - 全体のデフォルト
2. `styles.body` - ボディ部分全体
3. `styles.column[key]` - 特定の列全体
4. `styles.row(data, index)` - 特定の行全体
5. `headers[].style` - 列のセル全体に適用（オブジェクト）または条件付き（関数）
6. `rows[].{key}.style` - セル単位の直接スタイル（最優先）

## 主要機能の詳細

### 1. ヘッダーセルのスタイル指定

ヘッダー行の特定のセルにスタイルを付けることができます：

```typescript
headers: [
  { 
    key: 'age', 
    label: { value: '年齢', style: { font: { bold: true, color: '#0000FF' } } }  // ヘッダーセルにスタイル
  },
  { 
    key: 'name', 
    label: '名前'  // シンプルな文字列
  }
]
```

**ルール:**
- `label` が文字列の場合 → そのままヘッダーテキストとして使用
- `label` が `{ value, style }` の場合 → ヘッダーセルにスタイルを適用

### 2. セルレベルのスタイル指定

```typescript
rows: [
  { 
    age: 18,  // シンプルな値
    name: { value: "Mary", style: { font: { bold: true } } }  // スタイル付き
  }
]
```

**ルール:**
- 値が `{ value, style }` オブジェクトの場合 → セルスタイルとして認識
- 値がプリミティブ（文字列、数値など）の場合 → そのまま値として使用
- 両方の形式を同じ `rows` 配列内で混在可能

### 2. 結合（merge）

```typescript
headers: [
  { 
    key: 'dept', 
    label: '部署', 
    merge: 'vertical'  // 同じ値のセルを縦方向に自動結合
  }
]
```

**動作:**
- `merge: 'vertical'` を指定した列は、連続する同じ値のセルが自動的に結合される
- 結合されたセルは中央揃えが推奨（`styles.column` で設定可能）

### 3. 罫線（borders）

#### 方法1: プリセット（シンプル）
```typescript
{
  borders: 'all'  // 'all' | 'outer' | 'header-body' | 'none'
}
```

#### 方法2: 詳細設定
```typescript
{
  styles: {
    all: {
      border: {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      }
    }
  }
}
```

#### 方法3: セル単位で罫線
```typescript
rows: [
  { 
    age: 18,
    name: { 
      value: "Mary", 
      style: { 
        border: { bottom: { style: 'thick', color: { argb: 'FFFF0000' } } } 
      } 
    }
  }
]
```

### 4. 列幅自動調整

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

## 型定義

### コア型定義 (`src/types.ts`)

```typescript
// セル値の型（プリミティブまたはスタイル付き）
export type CellValue = any | {
  value: any;
  style?: XLStyle;
};

export interface HeaderDef {
  key: string;
  label: string | { value: string; style?: XLStyle };  // 文字列または{ value, style }形式
  width?: number | 'auto';
  merge?: 'vertical';
  style?: XLStyle | ((val: any, row: any, index: number) => XLStyle);  // オブジェクト（固定）または関数（条件付き）
  format?: string | ((val: any) => string);
}

export interface TitleConfig {
  label: string | string[];
  style?: XLStyle;
}

export interface StylesConfig {
  all?: XLStyle;  // 全体のデフォルト
  header?: XLStyle;  // ヘッダー行全体
  body?: XLStyle;  // ボディ全体
  row?: (data: any, index: number) => XLStyle;  // 行全体（動的）
  column?: { [key: string]: XLStyle };  // 列全体
}

export interface SheetConfig {
  name: string;
  headers: HeaderDef[];
  rows: any[];  // CellValue を含むデータオブジェクトの配列
  title?: TitleConfig;
  styles?: StylesConfig;  // 全体スタイル設定
  borders?: 'all' | 'outer' | 'header-body' | 'none';
  autoWidth?: boolean | {
    enabled?: boolean;
    padding?: number;
    headerIncluded?: boolean;
    charWidthConstant?: number;
  };
}
```

## 実装の主要ポイント

### 1. セル値の判定ロジック

```typescript
function isCellValueWithStyle(val: any): val is { value: any; style: XLStyle } {
  return val !== null && 
         typeof val === 'object' && 
         'value' in val && 
         !Array.isArray(val) &&
         !(val instanceof Date);
}

// 使用例
rows.forEach(row => {
  headers.forEach(header => {
    const cellData = row[header.key];
    if (isCellValueWithStyle(cellData)) {
      // cellData.value を値として使用
      // cellData.style をスタイルとして適用
    } else {
      // cellData をそのまま値として使用
    }
  });
});
```

### 2. スタイルの適用順序

```typescript
function getCellStyle(
  cellData: any,
  rowData: any,
  rowIndex: number,
  header: HeaderDef,
  styles: StylesConfig
): XLStyle {
  let finalStyle: XLStyle = {};
  
  // 1. styles.all
  if (styles.all) finalStyle = { ...finalStyle, ...styles.all };
  
  // 2. styles.body (ヘッダー行以外)
  if (styles.body && rowIndex > 0) finalStyle = { ...finalStyle, ...styles.body };
  
  // 3. styles.column[key]
  if (styles.column?.[header.key]) {
    finalStyle = { ...finalStyle, ...styles.column[header.key] };
  }
  
  // 4. styles.row(data, index)
  if (styles.row) {
    const rowStyle = styles.row(rowData, rowIndex);
    finalStyle = { ...finalStyle, ...rowStyle };
  }
  
  // 5. headers[].style（オブジェクトまたは関数）
  if (header.style) {
    if (typeof header.style === 'function') {
      // 条件付きスタイル（関数）
      const cellStyle = header.style(cellData, rowData, rowIndex);
      finalStyle = { ...finalStyle, ...cellStyle };
    } else {
      // 固定スタイル（オブジェクト）
      finalStyle = { ...finalStyle, ...header.style };
    }
  }
  
  // 6. セル単位のスタイル（最優先）
  if (isCellValueWithStyle(cellData) && cellData.style) {
    finalStyle = { ...finalStyle, ...cellData.style };
  }
  
  return finalStyle;
}
```

## 変更ファイル一覧

### コア実装
- [MODIFY] `src/types.ts` - 新しい型定義
- [MODIFY] `src/Sheetflow.ts` - 新しいAPIに対応
- [MODIFY] `src/index.ts` - export更新

### テストファイル（10ファイル）
- [MODIFY] `test/fill.test.ts`
- [MODIFY] `test/visual.test.ts`
- [MODIFY] `test/validation.test.ts`
- [MODIFY] `test/timeout.test.ts`
- [MODIFY] `test/performance.test.ts`
- [MODIFY] `test/multi_sheet.test.ts`
- [MODIFY] `test/layout.test.ts`
- [MODIFY] `test/format.test.ts`
- [MODIFY] `test/font.test.ts`
- [MODIFY] `test/border.test.ts`

### ドキュメント
- [MODIFY] `examples/demo.ts`
- [MODIFY] `README.md`

## 検証計画

### 自動テスト
```powershell
npm test
```

### 手動検証
1. デモサンプルの実行
2. 生成されたExcelファイルの確認
3. TypeScriptコンパイルの確認

### バージョンアップ
```json
{
  "version": "1.1.0"
}
```
