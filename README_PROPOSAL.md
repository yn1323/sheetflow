# README改善案

現在のREADMEに不足している重要な情報と、開発者が知っておくべきパターンをまとめました。これらをREADMEに追加することを推奨します。

## 1. 重要な制約事項 (Constraints)

### 予約語 (`Reserved Keywords`)
データ行のスタイル定義に `style` プロパティを使用するため、列のキー (`key`) として `'style'` を使用することはできません。使用すると実行時にエラーが発生します。

**追加案:**
```markdown
### ⚠️ 制約事項

- **予約語**: 列の `key` に `'style'` は使用できません。これは行ごとのスタイル定義プロパティとして予約されています。
- **シート名**: Excelの仕様により、シート名は31文字以内で、`\ / ? * [ ] :` の文字は使用できません。
```

### スタイルの優先順位 (`Style Precedence`)
スタイルが競合した場合の適用順序を明記する必要があります。

**追加案:**
```markdown
### スタイルの優先順位

スタイルは以下の順序で適用（上書き）されます（下が優先）：

1. **行定義 (`rows.style`)**: シート定義で指定した行全体のスタイル
2. **データ行スタイル (`row.style`)**: データオブジェクトに含まれる `style` プロパティ
3. **列定義 (`columns.style`)**: 列ごとのスタイル
4. **条件付き列スタイル**: `columns.style` が関数の場合
```

## 2. 推奨パターン (Recommended Patterns)

### 複数シートの作成
メソッドチェーンを使用して、1つのワークブックに複数のシートを追加するパターンを紹介します。

**追加案:**
```typescript
// 複数シートの作成
await createWorkbook()
  .addSheet(usersSheet, usersData)
  .addSheet(productsSheet, productsData)
  .save('data.xlsx');
```

### 型安全性 (`Type Safety`)
TypeScriptを使用する場合のメリットを強調します。

**追加案:**
```typescript
interface User {
  id: number;
  name: string;
}

// keyは 'id' | 'name' に型推論され、誤字を防げます
const sheet = defineSheet<User>({
  name: 'Users',
  columns: [
    { key: 'id', ... },   // OK
    { key: 'nmae', ... }  // Error: Type '"nmae"' is not assignable to type 'keyof User'
  ]
});
```

## 3. その他 (Others)

### `autoWidth` の詳細設定
`autoWidth` オプションでパディングや計算定数を調整できることも記載しておくと親切です。

**追加案:**
```typescript
const sheet = defineSheet<User>({
  // ...
  autoWidth: {
    padding: 4,            // デフォルト: 2
    headerIncluded: true,  // ヘッダーの長さも考慮するか
    charWidthConstant: 1.5 // 文字幅の係数（デフォルト: 1.2）
  }
});
```
