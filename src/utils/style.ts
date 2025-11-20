import { Style, Borders } from 'exceljs';
import { XLStyle } from '../types';
import { toArgb } from './color';

export function mapStyle(style: XLStyle): Partial<Style> {
  const excelStyle: Partial<Style> = {};

  if (style.font) {
    excelStyle.font = { ...style.font };
    if (typeof style.font.color === 'string') {
      excelStyle.font.color = { argb: toArgb(style.font.color) };
    }
  }

  if (style.fill) {
    if (style.fill.color) {
      excelStyle.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: toArgb(style.fill.color) },
      };
    } else {
      // If no color is specified but fill is provided, use it as-is
      excelStyle.fill = style.fill as any;
    }
  }

  if (style.alignment) {
    excelStyle.alignment = style.alignment;
  }

  if (style.border) {
    if (typeof style.border === 'string') {
      // Presets handled in main logic or expanded here
      // For now, we'll handle explicit border objects here
    } else {
      excelStyle.border = style.border as Borders;
    }
  }

  return excelStyle;
}
