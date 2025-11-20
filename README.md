# xlkit

<p align="center">
  <img src="./logo.png" alt="xlkit Logo" width="200" />
</p>

[ExcelJS](https://github.com/exceljs/exceljs) のための宣言的スキーマベースラッパーです。
シンプルなスキーマでExcelの構造を定義するだけで、スタイル、フォーマット、レイアウトをxlkitが自動で処理します。

[English README](./README_en.md)

## 特徴

- 📝 **宣言的スキーマ**: 列、ヘッダー、スタイルを一箇所で定義。
- 🎨 **高度なスタイル設定**: ヘッダー、行、列、または特定のセルに対して、条件付きでスタイルを適用可能。
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
import { createWorkbook, defineSheet } from 'xlkit';

// 1. データ型を定義
interface User {
  id: number;
  name: string;
  role: string;
  isActive: boolean;
}

// 2. シートのスキーマを定義
const userSheet = defineSheet<User>({
  name: 'Users',
  columns: [
    { key: 'id', header: 'ID', width: 10 },
    { key: 'name', header: '氏名', width: 20 },
    { key: 'role', header: '役割', width: 'auto', merge: 'vertical' },
    { 
      key: 'isActive', 
      header: 'ステータス', 
      format: (val) => val ? '有効' : '無効',
      style: (val) => ({ font: { color: val ? '#00AA00' : '#FF0000' } })
    }
  ],
  borders: 'outer'
});

// 3. Excel生成
await createWorkbook().addSheet(userSheet, users).save('users.xlsx');
```

## 詳細リファレンス

### 1. レイアウトと列定義 (`columns`)

列の定義は `columns` 配列で行います。

| プロパティ | 型 | 説明 |
| :--- | :--- | :--- |
| `key` | `keyof T` | データのプロパティキー |
| `header` | `string` | ヘッダーに表示するテキスト |
| `width` | `number` \| `'auto'` | 列幅。`'auto'` を指定するとコンテンツに合わせて自動調整されます。 |
| `merge` | `'vertical'` | `'vertical'` を指定すると、上下のセルが同じ値の場合に自動結合されます。 |
| `format` | `string` \| `Function` | 数値/日付のフォーマット文字列（例: `'$#,##0'`, `'yyyy-mm-dd'`）または変換関数。 |
| `style` | `Style` \| `Function` | 列全体のスタイル、または値に応じた条件付きスタイル関数。 |

### 2. 罫線 (`borders`)

シート全体の罫線プリセットを `borders` プロパティで指定できます。

- **`'all'`**: すべてのセルに格子状の罫線を引きます。
- **`'outer'`**: データ領域の外枠のみに罫線を引きます。
- **`'header-body'`**: ヘッダー行の下に太めの線を引きます。
- **`'none'`**: 罫線を引きません（デフォルト）。

```typescript
const sheet = defineSheet<User>({
  // ...
  borders: 'all', // 格子状
});
```

個別のセルに罫線を設定したい場合は、`style` プロパティ内の `border` で詳細に指定可能です。

### 3. スタイル (`style`)

フォント、背景色、配置などを詳細に設定できます。6桁のHexカラーコード (`#RRGGBB`) が使用可能です。

```typescript
style: {
  font: { 
    name: 'Arial',
    size: 12,
    bold: true,
    color: '#FF0000' // 赤色
  },
  fill: {
    color: '#EEEEEE' // Hexカラーで簡単に指定
  },
  alignment: {
    vertical: 'middle',
    horizontal: 'center',
    wrapText: true
  },
  border: {
    top: { style: 'thin', color: { argb: 'FF000000' } },
    // ...
  }
}
```

### 4. ヘッダー設定 (`header`)

ヘッダー行のスタイルや構成をカスタマイズできます。

```typescript
header: {
  rows: ['社員名簿 2025'], // 1行目（タイトルなど）
  style: { // ヘッダー行のスタイル
    font: { bold: true, color: '#FFFFFF' },
    fill: { color: '#4472C4' }
  },
  borders: 'header-body' // ヘッダー下の罫線
}
```

### 5. 行スタイル (`rows`)

行ごとのスタイル（縞模様など）を定義できます。

```typescript
rows: {
  style: (data, index) => {
    // 偶数行に背景色をつける
    return index % 2 === 0 ? { fill: { color: '#F0F0F0' } } : {};
  }
}
```

### 6. ブラウザ環境でのダウンロード

ブラウザ環境では、`download()` メソッドを使って簡単にExcelファイルをダウンロードできます。

```typescript
// Node.js環境
await createWorkbook().addSheet(sheet, data).save('output.xlsx');

// ブラウザ環境
await createWorkbook().addSheet(sheet, data).download('output.xlsx');
```

内部的には`saveToBuffer()`を呼び出してBlobを作成し、自動的にダウンロードを開始します。

### 7. タイムアウト設定

大量データ処理時のフリーズを防ぐため、`save()`、`saveToBuffer()`、`download()` にはデフォルトで10秒のタイムアウトが設定されています。

```typescript
// デフォルト（10秒）
await createWorkbook().addSheet(sheet, data).save('output.xlsx');

// カスタムタイムアウト（30秒）
await createWorkbook().addSheet(sheet, data).save('output.xlsx', { timeout: 30000 });

// ブラウザ環境でも同様
await createWorkbook().addSheet(sheet, data).download('output.xlsx', { timeout: 15000 });
```

> **推奨**: 10万行以下のデータであればデフォルト設定で問題ありません。それ以上の大量データを扱う場合は、ファイル分割やストリーミング処理を検討してください。

## ライセンス

MIT
