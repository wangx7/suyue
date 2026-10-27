import { read, utils, type WorkBook, type WorkSheet } from 'xlsx';
import type { CellRange, SheetData, SheetWorkerResponse } from '../renderers/sheet-types';

// 解析放在 Worker 中，大表格也不会卡住界面
const ctx = self as unknown as {
  onmessage: (e: MessageEvent<{ buf: ArrayBuffer; name: string }>) => void;
  postMessage(msg: SheetWorkerResponse): void;
};

/** 中文 CSV 常见 GBK 编码：先按 UTF-8 严格解码，失败再退回 GB18030 */
function decodeText(buf: ArrayBuffer): string {
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(buf);
  } catch {
    return new TextDecoder('gb18030').decode(buf);
  }
}

/** 实际有内容的范围。很多文件的 !ref 因格式设置被撑到上百万行，不能直接用 */
function usedExtent(ws: WorkSheet, merges: CellRange[]) {
  const data = (ws['!data'] ?? []) as ({ v?: unknown; w?: string } | undefined)[][];
  let maxR = -1;
  let maxC = -1;
  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    if (!row) continue;
    for (let c = row.length - 1; c >= 0; c--) {
      const cell = row[c];
      if (cell && ((cell.v != null && cell.v !== '') || cell.w)) {
        if (c > maxC) maxC = c;
        maxR = r;
        break;
      }
    }
  }
  for (const m of merges) {
    if (m.e.r > maxR) maxR = m.e.r;
    if (m.e.c > maxC) maxC = m.e.c;
  }
  return { maxR, maxC };
}

function toSheet(name: string, ws: WorkSheet): SheetData {
  const merges = (ws['!merges'] ?? []) as CellRange[];
  const { maxR, maxC } = usedExtent(ws, merges);
  const rows =
    maxR < 0
      ? []
      : utils.sheet_to_json<string[]>(ws, {
          header: 1,
          raw: false,
          blankrows: true,
          range: { s: { r: 0, c: 0 }, e: { r: maxR, c: maxC } },
        });

  const colWidths: (number | null)[] = [];
  (ws['!cols'] ?? []).slice(0, maxC + 1).forEach((col, i) => {
    if (!col) return;
    if (col.hidden) colWidths[i] = 0;
    else if (col.wpx) colWidths[i] = col.wpx;
    else if (col.wch) colWidths[i] = Math.round(col.wch * 7 + 5);
    else if (col.width) colWidths[i] = Math.round(col.width * 7 + 5);
  });
  const rowHeights: (number | null)[] = [];
  (ws['!rows'] ?? []).slice(0, maxR + 1).forEach((row, i) => {
    if (!row) return;
    if (row.hidden) rowHeights[i] = 0;
    else if (row.hpx) rowHeights[i] = row.hpx;
    else if (row.hpt) rowHeights[i] = Math.round((row.hpt * 4) / 3);
  });
  return { name, rows, merges, colWidths, rowHeights };
}

ctx.onmessage = (e) => {
  const { buf, name } = e.data;
  try {
    const ext = /\.([a-z0-9]+)$/i.exec(name)?.[1].toLowerCase() ?? '';
    const wb: WorkBook = ['csv', 'tsv', 'txt'].includes(ext)
      ? read(decodeText(buf), { type: 'string', dense: true, raw: true }) // raw：保留 00123、长数字等原文
      : read(buf, { type: 'array', dense: true, cellDates: true, cellStyles: true, dateNF: 'yyyy/m/d' }); // 默认短日期按中文习惯显示

    const meta = wb.Workbook?.Sheets ?? [];
    let names = wb.SheetNames.filter((_, i) => !meta[i]?.Hidden);
    if (!names.length) names = wb.SheetNames;
    ctx.postMessage({ ok: true, sheets: names.map((n) => toSheet(n, wb.Sheets[n])) });
  } catch (err) {
    ctx.postMessage({ ok: false, error: err instanceof Error ? err.message : String(err) });
  }
};
