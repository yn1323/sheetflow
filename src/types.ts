import { Style, Alignment, Border, Fill, Font } from 'exceljs';

export type HexColor = string; // #RRGGBB

export interface XLStyle {
  font?: Partial<Font> & { color?: HexColor | { argb: string } };
  fill?: Partial<Fill> & { color?: HexColor };
  alignment?: Partial<Alignment>;
  border?: Partial<Border> | 'all' | 'outer' | 'header-body' | 'none';
}

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
