export interface CellRange {
  s: { r: number; c: number };
  e: { r: number; c: number };
}

export interface SheetData {
  name: string;
  /** 格式化后的文本，稀疏数组（空单元格为空洞） */
  rows: string[][];
  merges: CellRange[];
  /** 像素宽度；null = 默认，0 = 隐藏 */
  colWidths: (number | null)[];
  rowHeights: (number | null)[];
}

export type SheetWorkerResponse = { ok: true; sheets: SheetData[] } | { ok: false; error: string };
