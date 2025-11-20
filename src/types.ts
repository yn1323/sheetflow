import { Style, Alignment, Border, Fill, Font } from 'exceljs';

export type HexColor = string; // #RRGGBB

export interface XLStyle {
  font?: Partial<Font> & { color?: HexColor | { argb: string } };
  fill?: Partial<Fill> & { color?: HexColor };
  alignment?: Partial<Alignment>;
  border?: Partial<Border> | 'all' | 'outer' | 'header-body' | 'none';
}

export interface ColumnDef<T> {
  key: keyof T;
  header: string;
  width?: number | 'auto';
  merge?: 'vertical';
  style?: XLStyle | ((val: any, row: T, index: number) => XLStyle);
  format?: string | ((val: any) => string);
}

export interface HeaderConfig {
  rows: string[];
  style?: XLStyle;
  borders?: 'header-body' | 'all' | 'none';
}

export interface SheetDef<T> {
  name: string;
  columns: ColumnDef<T>[];
  header?: HeaderConfig;
  rows?: {
    style?: (data: T, index: number) => XLStyle;
  };
  defaultStyle?: XLStyle;
  borders?: 'all' | 'outer' | 'header-body' | 'none';
  autoWidth?: {
    padding?: number;
    headerIncluded?: boolean;
    charWidthConstant?: number;
  };
}
