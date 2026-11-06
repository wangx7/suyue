import { icon } from '../icons';
import { cssVar, el } from '../ui';
import { showFilterMenu } from './sheet-filter';
import type { CellRange, SheetData, SheetWorkerResponse } from './sheet-types';
import workerUrl from '../workers/sheet.worker.ts?worker&url';
import type { Renderer } from './types';

const ROW_H = 24;
const HEAD_H = 24;
const COL_W = 96;
const PAD = 6;
const LINE_H = 17;
const FONT = '13px -apple-system, BlinkMacSystemFont, "PingFang SC", "Microsoft YaHei", sans-serif';
const HEAD_FONT = '12px -apple-system, BlinkMacSystemFont, "PingFang SC", "Microsoft YaHei", sans-serif';
const NUM_RE = /^[-+(]?[¥$€£]?\s?[\d,]*\.?\d+(?:[eE][-+]?\d+)?%?\)?$/;
/** 数字右对齐；但前导零（00123）和超过 15 位的数字串是编号，按文本左对齐 */
const isNumeric = (t: string) => NUM_RE.test(t) && !/^0\d/.test(t) && !/\d{16}/.test(t);
/** 数据区外额外显示的空白行列，观感接近 Excel */
const EXTRA_ROWS = 30;
const EXTRA_COLS = 5;
/** 列头右侧筛选按钮的宽度 */
const FUN_W = 20;
const collator = new Intl.Collator('zh-CN', { numeric: true });

/** 一张表的筛选状态：header 行以下的行按各列允许的值显示 */
interface Filter {
  header: number;
  cols: Map<number, Set<string>>;
}

/** 单元格里有制表符、换行或引号时按 Excel 的规则加引号，粘贴回 Excel 时才不会错位 */
const quoteCell = (t: string) => (/[\t\r\n"]/.test(t) ? `"${t.replace(/"/g, '""')}"` : t);

async function writeClipboard(text: string) {
  try {
    await navigator.clipboard.writeText(text);
    return;
  } catch {
    /* 没有剪贴板权限时退回 execCommand */
  }
  const area = el('textarea');
  area.value = text;
  area.style.cssText = 'position:fixed;left:-9999px;top:0;opacity:0';
  document.body.append(area);
  area.select();
  document.execCommand('copy');
  area.remove();
}

async function createWorker(): Promise<Worker> {
  try {
    return new Worker(workerUrl, { type: 'module' });
  } catch {
    // VS Code 的 webview 里脚本和页面不同源，不能直接建 Worker；Worker 打包成了单文件，取回来用 blob 地址建
    const code = await (await fetch(workerUrl)).text();
    return new Worker(URL.createObjectURL(new Blob([code], { type: 'text/javascript' })), { type: 'module' });
  }
}

async function parseInWorker(buf: ArrayBuffer, name: string): Promise<SheetData[]> {
  const worker = await createWorker();
  return new Promise((resolve, reject) => {
    worker.onmessage = (e: MessageEvent<SheetWorkerResponse>) => {
      worker.terminate();
      if (e.data.ok) resolve(e.data.sheets);
      else reject(new Error(e.data.error));
    };
    worker.onerror = (e) => {
      worker.terminate();
      reject(new Error(e.message || '表格解析失败'));
    };
    worker.postMessage({ buf, name }, [buf]);
  });
}

function colName(c: number): string {
  let s = '';
  for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

function prefixSums(n: number, size: (i: number) => number): Float64Array {
  const a = new Float64Array(n + 1);
  for (let i = 0; i < n; i++) a[i + 1] = a[i] + size(i);
  return a;
}

/** 满足 pre[i] <= x 的最大 i */
function indexAt(pre: Float64Array, x: number): number {
  let lo = 0;
  let hi = pre.length - 2;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (pre[mid] <= x) lo = mid;
    else hi = mid - 1;
  }
  return lo;
}

/** Canvas 虚拟网格：只绘制可视区域，几十万行也能流畅滚动 */
class Grid {
  private canvas = el('canvas');
  private ctx = this.canvas.getContext('2d')!;
  private spacer = el('div');
  private raf = 0;
  private colors: Record<string, string> = {};

  private sheet!: SheetData;
  private nRows = 0;
  private nCols = 0;
  private colX: Float64Array = new Float64Array(1);
  private rowY: Float64Array = new Float64Array(1);
  private rowNumW = 40;
  /** 被合并覆盖的单元格 → 所属合并区域 */
  private covered = new Map<number, CellRange>();
  /** 当前单元格（选区的起点）和选区的另一端 */
  private sel = { r: 0, c: 0 };
  private end = { r: 0, c: 0 };
  private hit = false;
  private drag: 'cell' | 'row' | 'col' | null = null;
  private pointer = { x: 0, y: 0 };
  private dragRaf = 0;
  private dataCols = 0;
  /** 被筛选隐藏的行 */
  private hidden = new Uint8Array(0);
  private filters = new Map<SheetData, Filter>();
  private f: Filter | null = null;
  /** 缩放系数：内部一律使用未缩放的逻辑坐标，只在绘制、滚动和鼠标换算时乘除 z */
  private z = 1;

  onSelect: (ref: string, value: string) => void = () => {};
  onCopy: (rows: number, cols: number) => void = () => {};
  /** x、y 为列头下方的位置（相对表格区域左上角的像素） */
  onFilterClick: (col: number, x: number, y: number) => void = () => {};
  onFilterChange: (visible: number, total: number, active: boolean) => void = () => {};

  constructor(private scroller: HTMLElement, gridBox: HTMLElement) {
    gridBox.append(this.canvas, scroller);
    scroller.append(this.spacer);
    scroller.tabIndex = 0;
    scroller.addEventListener('scroll', this.schedule);
    scroller.addEventListener('mousedown', this.onMouseDown);
    scroller.addEventListener('keydown', this.onKeyDown);
    scroller.addEventListener('copy', this.onCopyEvent);
    this.readColors();
  }

  readColors() {
    for (const k of ['grid-bg', 'grid-line', 'grid-head', 'grid-head-text', 'grid-text', 'accent', 'accent-soft', 'grid-hit', 'grid-range']) {
      this.colors[k] = cssVar(`--${k}`);
    }
    this.schedule();
  }

  setSheet(sheet: SheetData) {
    this.sheet = sheet;
    this.dataCols = sheet.rows.reduce((m, row) => Math.max(m, row.length), 0);
    this.nRows = sheet.rows.length + EXTRA_ROWS;
    this.nCols = Math.max(this.dataCols + EXTRA_COLS, 26);
    this.colX = prefixSums(this.nCols, (c) => sheet.colWidths[c] ?? COL_W);
    this.f = this.filters.get(sheet) ?? null;
    this.buildRows();
    this.ctx.font = HEAD_FONT;
    this.rowNumW = Math.max(40, Math.ceil(this.ctx.measureText(String(this.nRows)).width) + 18);

    this.covered.clear();
    for (const m of sheet.merges) {
      if ((m.e.r - m.s.r + 1) * (m.e.c - m.s.c + 1) > 1e5) continue;
      for (let r = m.s.r; r <= m.e.r; r++) for (let c = m.s.c; c <= m.e.c; c++) this.covered.set(this.key(r, c), m);
    }

    this.sizeSpacer();
    this.scroller.scrollTo(0, 0);
    this.select(0, 0);
    this.reportFilter();
  }

  // ---------- 筛选 ----------
  private buildRows() {
    const s = this.sheet;
    this.hidden = new Uint8Array(s.rows.length);
    if (this.f) for (let r = this.f.header + 1; r < s.rows.length; r++) if (!this.passes(r, -1)) this.hidden[r] = 1;
    this.rowY = prefixSums(this.nRows, (r) => (this.hidden[r] ? 0 : (s.rowHeights[r] ?? ROW_H)));
  }

  /** 第 r 行是否满足除 except 列以外的所有筛选条件 */
  private passes(r: number, except: number): boolean {
    const row = this.sheet.rows[r];
    for (const [c, allowed] of this.f!.cols) if (c !== except && !allowed.has(row?.[c] ?? '')) return false;
    return true;
  }

  private reportFilter() {
    const f = this.f;
    const total = f ? Math.max(0, this.sheet.rows.length - 1 - f.header) : 0;
    let shown = total;
    for (const h of this.hidden) shown -= h;
    this.onFilterChange(shown, total, !!f && f.cols.size > 0);
  }

  private applyFilter() {
    this.buildRows();
    this.sizeSpacer();
    this.reportFilter();
    this.schedule();
  }

  get filtering() {
    return !!this.f;
  }

  /** 开启筛选：以当前选中的行作为表头；再次调用则关闭并显示全部行 */
  toggleFilter(): boolean {
    if (this.f) {
      this.filters.delete(this.sheet);
      this.f = null;
    } else {
      if (!this.sheet.rows.length) return false;
      this.f = { header: Math.min(this.sel.r, this.sheet.rows.length - 1), cols: new Map() };
      this.filters.set(this.sheet, this.f);
    }
    this.applyFilter();
    return !!this.f;
  }

  clearFilters() {
    if (!this.f?.cols.size) return;
    this.f.cols.clear();
    this.applyFilter();
  }

  /** 第 c 列可选的值：只统计满足其他列筛选条件的行，和 Excel 一样逐层缩小 */
  columnValues(c: number): { values: [string, number][]; allowed: Set<string> | null } {
    const f = this.f!;
    const counts = new Map<string, number>();
    for (let r = f.header + 1; r < this.sheet.rows.length; r++) {
      if (!this.passes(r, c)) continue;
      const v = this.sheet.rows[r]?.[c] ?? '';
      counts.set(v, (counts.get(v) ?? 0) + 1);
    }
    const values = [...counts].sort((a, b) => (a[0] === '' ? -1 : b[0] === '' ? 1 : collator.compare(a[0], b[0])));
    return { values, allowed: f.cols.get(c) ?? null };
  }

  setColumnFilter(c: number, allowed: Set<string> | null) {
    if (!this.f) return;
    if (allowed) this.f.cols.set(c, allowed);
    else this.f.cols.delete(c);
    this.applyFilter();
  }

  private sizeSpacer() {
    this.spacer.style.width = `${(this.rowNumW + this.colX[this.nCols]) * this.z}px`;
    this.spacer.style.height = `${(HEAD_H + this.rowY[this.nRows]) * this.z}px`;
  }

  get zoom() {
    return this.z;
  }

  /** 以 anchor（相对表格左上角的像素坐标）为中心缩放，指针下的单元格保持不动 */
  setZoom(z: number, anchor = { x: 0, y: 0 }) {
    const s = this.scroller;
    const lx = (s.scrollLeft + anchor.x) / this.z;
    const ly = (s.scrollTop + anchor.y) / this.z;
    this.z = z;
    this.sizeSpacer();
    s.scrollLeft = lx * z - anchor.x;
    s.scrollTop = ly * z - anchor.y;
    this.schedule();
  }

  get selection(): [number, number] {
    return [this.sel.r, this.sel.c];
  }

  private key(r: number, c: number) {
    return r * 16384 + c;
  }

  private text(r: number, c: number): string {
    return this.sheet.rows[r]?.[c] ?? '';
  }

  schedule = () => {
    if (!this.raf) this.raf = requestAnimationFrame(this.draw);
  };

  destroy() {
    cancelAnimationFrame(this.raf);
    cancelAnimationFrame(this.dragRaf);
    this.onMouseUp();
  }

  // ---------- 选择与导航 ----------
  private clampCell(r: number, c: number) {
    return { r: Math.max(0, Math.min(r, this.nRows - 1)), c: Math.max(0, Math.min(c, this.nCols - 1)) };
  }

  select(r: number, c: number, fromSearch = false) {
    ({ r, c } = this.clampCell(r, c));
    const m = this.covered.get(this.key(r, c));
    if (m) ({ r, c } = m.s);
    this.sel = { r, c };
    this.end = { r, c };
    this.hit = fromSearch;
    this.notify();
    this.ensureVisible(r, c);
    this.schedule();
  }

  /** 把选区的另一端移到 (r, c)，起点不变 */
  extend(r: number, c: number) {
    this.end = this.clampCell(r, c);
    this.hit = false;
    this.notify();
    this.ensureVisible(this.end.r, this.end.c);
    this.schedule();
  }

  private setRange(from: { r: number; c: number }, to: { r: number; c: number }) {
    this.sel = from;
    this.end = to;
    this.hit = false;
    this.notify();
    this.schedule();
  }

  selectAll() {
    this.setRange({ r: 0, c: 0 }, { r: Math.max(0, this.sheet.rows.length - 1), c: Math.max(0, this.dataCols - 1) });
  }

  private bounds() {
    const { sel, end } = this;
    return { r0: Math.min(sel.r, end.r), r1: Math.max(sel.r, end.r), c0: Math.min(sel.c, end.c), c1: Math.max(sel.c, end.c) };
  }

  private notify() {
    const { r0, r1, c0, c1 } = this.bounds();
    const cell = (r: number, c: number) => `${colName(c)}${r + 1}`;
    let ref = cell(r0, c0);
    if (r0 !== r1 || c0 !== c1) {
      if (r0 === 0 && r1 === this.nRows - 1) ref = `${colName(c0)}:${colName(c1)}`;
      else if (c0 === 0 && c1 === this.nCols - 1) ref = `${r0 + 1}:${r1 + 1}`;
      else ref = `${ref}:${cell(r1, c1)}`;
    }
    this.onSelect(ref, this.text(this.sel.r, this.sel.c));
  }

  // ---------- 复制 ----------
  /** 选区文本：制表符分隔列、换行分隔行，跳过被隐藏或筛掉的行列，范围超出数据时只取有数据的部分 */
  private selectionText(): { text: string; rows: number; cols: number } {
    const { r0, c0 } = this.bounds();
    let { r1, c1 } = this.bounds();
    if (r0 === r1 && c0 === c1) return { text: this.text(r0, c0), rows: 1, cols: 1 };
    r1 = Math.min(r1, this.sheet.rows.length - 1);
    c1 = Math.min(c1, this.dataCols - 1);
    const cols: number[] = [];
    for (let c = c0; c <= c1; c++) if (this.colX[c + 1] > this.colX[c]) cols.push(c);
    const lines: string[] = [];
    for (let r = r0; r <= r1; r++) {
      if (this.rowY[r + 1] === this.rowY[r]) continue;
      lines.push(cols.map((c) => quoteCell(this.text(r, c))).join('\t'));
    }
    return { text: lines.join('\n'), rows: lines.length, cols: cols.length };
  }

  async copy() {
    const { text, rows, cols } = this.selectionText();
    if (!rows || !cols) return;
    await writeClipboard(text);
    this.onCopy(rows, cols);
  }

  // 浏览器菜单里的“复制”走这里；快捷键由 onKeyDown 处理
  private onCopyEvent = (e: ClipboardEvent) => {
    const { text, rows, cols } = this.selectionText();
    if (!rows || !cols) return;
    e.clipboardData?.setData('text/plain', text);
    e.preventDefault();
    this.onCopy(rows, cols);
  };

  private ensureVisible(r: number, c: number) {
    const s = this.scroller;
    const z = this.z;
    const viewW = s.clientWidth / z - this.rowNumW;
    const viewH = s.clientHeight / z - HEAD_H;
    const sx = s.scrollLeft / z;
    const sy = s.scrollTop / z;
    const m = this.covered.get(this.key(r, c));
    const [top, bottom] = [this.rowY[r], this.rowY[(m?.e.r ?? r) + 1]];
    const [left, right] = [this.colX[c], this.colX[(m?.e.c ?? c) + 1]];
    if (top < sy) s.scrollTop = top * z;
    else if (bottom > sy + viewH) s.scrollTop = Math.min(top, bottom - viewH) * z;
    if (left < sx) s.scrollLeft = left * z;
    else if (right > sx + viewW) s.scrollLeft = Math.min(left, right - viewW) * z;
  }

  // ---------- 鼠标 ----------
  /** 指针下的区域和行列；clamp 为真时把指针限制在单元格区域内（拖选出界时用） */
  private locate(clientX: number, clientY: number, clamp = false) {
    const s = this.scroller;
    const rect = s.getBoundingClientRect();
    let px = (clientX - rect.left) / this.z;
    let py = (clientY - rect.top) / this.z;
    if (clamp) {
      px = Math.min(Math.max(px, this.rowNumW), s.clientWidth / this.z - 1);
      py = Math.min(Math.max(py, HEAD_H), s.clientHeight / this.z - 1);
    }
    const zone: 'corner' | 'row' | 'col' | 'cell' = px < this.rowNumW ? (py < HEAD_H ? 'corner' : 'row') : py < HEAD_H ? 'col' : 'cell';
    const c = indexAt(this.colX, Math.max(0, px - this.rowNumW) + s.scrollLeft / this.z);
    const r = indexAt(this.rowY, Math.max(0, py - HEAD_H) + s.scrollTop / this.z);
    return { zone, r, c, px };
  }

  private onMouseDown = (e: MouseEvent) => {
    if (e.button !== 0) return;
    const rect = this.scroller.getBoundingClientRect();
    if (e.clientX - rect.left > this.scroller.clientWidth || e.clientY - rect.top > this.scroller.clientHeight) return; // 点在滚动条上
    const { zone, r, c, px } = this.locate(e.clientX, e.clientY);
    e.preventDefault(); // 不让拖动变成页面文字选择
    this.scroller.focus({ preventScroll: true });

    if (zone === 'col' && this.f && c < this.dataCols && this.colX[c + 1] > this.colX[c]) {
      const sx = this.scroller.scrollLeft / this.z;
      if (px >= this.rowNumW + this.colX[c + 1] - sx - FUN_W) {
        this.onFilterClick(c, (this.rowNumW + this.colX[c] - sx) * this.z, (HEAD_H + 2) * this.z);
        return;
      }
    }
    if (zone === 'corner') return this.selectAll();
    if (zone === 'col') this.setRange({ r: 0, c }, { r: this.nRows - 1, c });
    else if (zone === 'row') this.setRange({ r, c: 0 }, { r, c: this.nCols - 1 });
    else if (e.shiftKey) this.extend(r, c);
    else this.select(r, c);
    this.drag = zone;
    this.pointer = { x: e.clientX, y: e.clientY };
    window.addEventListener('mousemove', this.onMouseMove);
    window.addEventListener('mouseup', this.onMouseUp);
  };

  private onMouseMove = (e: MouseEvent) => {
    this.pointer = { x: e.clientX, y: e.clientY };
    if (!this.dragRaf) this.dragRaf = requestAnimationFrame(this.dragTick);
  };

  private onMouseUp = () => {
    this.drag = null;
    window.removeEventListener('mousemove', this.onMouseMove);
    window.removeEventListener('mouseup', this.onMouseUp);
  };

  /** 拖选：按指针位置延伸选区；指针拖到表格外时自动滚动 */
  private dragTick = () => {
    this.dragRaf = 0;
    if (!this.drag) return;
    const s = this.scroller;
    const rect = s.getBoundingClientRect();
    const over = (v: number, lo: number, hi: number) => (v < lo ? v - lo : v > hi ? v - hi : 0);
    const dx = this.drag === 'row' ? 0 : over(this.pointer.x, rect.left + this.rowNumW * this.z, rect.left + s.clientWidth);
    const dy = this.drag === 'col' ? 0 : over(this.pointer.y, rect.top + HEAD_H * this.z, rect.top + s.clientHeight);
    s.scrollLeft += Math.max(-40, Math.min(40, dx));
    s.scrollTop += Math.max(-40, Math.min(40, dy));
    const { r, c } = this.locate(this.pointer.x, this.pointer.y, true);
    if (this.drag === 'col') this.setRange(this.sel, { r: this.nRows - 1, c });
    else if (this.drag === 'row') this.setRange(this.sel, { r, c: this.nCols - 1 });
    else {
      this.end = { r, c };
      this.notify();
      this.schedule();
    }
    if (dx || dy) this.dragRaf = requestAnimationFrame(this.dragTick);
  };

  // ---------- 键盘 ----------
  private step(pre: Float64Array, from: number, dir: number, max: number) {
    let i = from + dir;
    while (i >= 0 && i < max && pre[i + 1] === pre[i]) i += dir; // 跳过隐藏行列
    return i < 0 || i >= max ? from : i;
  }

  private onKeyDown = (e: KeyboardEvent) => {
    const { r, c } = this.sel;
    const m = this.covered.get(this.key(r, c));
    const endR = m?.e.r ?? r;
    const endC = m?.e.c ?? c;
    const mod = e.metaKey || e.ctrlKey;
    if (mod && e.key.toLowerCase() === 'c') {
      e.preventDefault();
      void this.copy();
      return;
    }
    if (mod && e.key.toLowerCase() === 'a') {
      e.preventDefault();
      this.selectAll();
      return;
    }
    if (e.key === 'Escape' && (this.end.r !== r || this.end.c !== c)) {
      this.select(r, c);
      return;
    }
    if (e.shiftKey && e.key.startsWith('Arrow')) {
      e.preventDefault();
      const { r: er, c: ec } = this.end;
      if (e.key === 'ArrowUp') this.extend(this.step(this.rowY, er, -1, this.nRows), ec);
      else if (e.key === 'ArrowDown') this.extend(this.step(this.rowY, er, 1, this.nRows), ec);
      else if (e.key === 'ArrowLeft') this.extend(er, this.step(this.colX, ec, -1, this.nCols));
      else this.extend(er, this.step(this.colX, ec, 1, this.nCols));
      return;
    }
    const moves: Record<string, () => void> = {
      ArrowUp: () => this.select(this.step(this.rowY, r, -1, this.nRows), c),
      ArrowDown: () => this.select(this.step(this.rowY, endR, 1, this.nRows), c),
      Enter: () => this.select(this.step(this.rowY, endR, 1, this.nRows), c),
      ArrowLeft: () => this.select(r, this.step(this.colX, c, -1, this.nCols)),
      ArrowRight: () => this.select(r, this.step(this.colX, endC, 1, this.nCols)),
      Tab: () => this.select(r, this.step(this.colX, endC, 1, this.nCols)),
    };
    const move = moves[e.key];
    if (move) {
      e.preventDefault();
      move();
    }
  };

  /** 从当前选中格往后（或往前）按行查找，循环到头 */
  find(query: string, backward: boolean): boolean {
    const q = query.trim().toLowerCase();
    if (!q) return false;
    const rows = this.sheet.rows;
    const cols = this.nCols;
    const total = rows.length * cols;
    if (!total) return false;
    let pos = this.sel.r * cols + this.sel.c;
    for (let i = 0; i < total; i++) {
      pos = (pos + (backward ? -1 : 1) + total) % total;
      const r = Math.floor(pos / cols);
      const row = rows[r];
      if (!row || !row.length || this.hidden[r]) {
        // 整行为空或被筛选隐藏时直接跳过这一行
        pos = backward ? r * cols : r * cols + cols - 1;
        continue;
      }
      const t = row[pos % cols];
      if (t && t.toLowerCase().includes(q)) {
        this.select(r, pos % cols, true);
        return true;
      }
    }
    return false;
  }

  // ---------- 绘制 ----------
  private draw = () => {
    this.raf = 0;
    if (!this.sheet) return;
    const { ctx, colors: k, colX, rowY, scroller } = this;
    const cw = scroller.clientWidth;
    const ch = scroller.clientHeight;
    const dpr = window.devicePixelRatio || 1;
    if (this.canvas.width !== Math.round(cw * dpr) || this.canvas.height !== Math.round(ch * dpr)) {
      this.canvas.width = Math.round(cw * dpr);
      this.canvas.height = Math.round(ch * dpr);
      this.canvas.style.width = `${cw}px`;
      this.canvas.style.height = `${ch}px`;
    }
    // 以下全部使用逻辑坐标，缩放交给画布变换
    const z = this.z;
    ctx.setTransform(dpr * z, 0, 0, dpr * z, 0, 0);
    const w = cw / z;
    const h = ch / z;
    const sx = scroller.scrollLeft / z;
    const sy = scroller.scrollTop / z;
    const RW = this.rowNumW;
    const X = (c: number) => RW + colX[c] - sx;
    const Y = (r: number) => HEAD_H + rowY[r] - sy;
    const r0 = indexAt(rowY, sy);
    const r1 = indexAt(rowY, sy + h - HEAD_H);
    const c0 = indexAt(colX, sx);
    const c1 = indexAt(colX, sx + w - RW);

    ctx.fillStyle = k['grid-bg'];
    ctx.fillRect(0, 0, w, h);

    ctx.save();
    ctx.beginPath();
    ctx.rect(RW, HEAD_H, w - RW, h - HEAD_H);
    ctx.clip();

    // 筛选的表头行用表头底色，和数据行区分开
    if (this.f && this.f.header >= r0 && this.f.header <= r1) {
      ctx.fillStyle = k['grid-head'];
      ctx.fillRect(RW, Y(this.f.header), w - RW, rowY[this.f.header + 1] - rowY[this.f.header]);
    }

    // 网格线
    ctx.strokeStyle = k['grid-line'];
    ctx.lineWidth = 1 / this.z;
    ctx.beginPath();
    for (let c = c0; c <= c1; c++) {
      const x = Math.round(X(c + 1)) - 0.5;
      ctx.moveTo(x, HEAD_H);
      ctx.lineTo(x, h);
    }
    for (let r = r0; r <= r1; r++) {
      const y = Math.round(Y(r + 1)) - 0.5;
      ctx.moveTo(RW, y);
      ctx.lineTo(w, y);
    }
    ctx.stroke();

    // 当前单元格和选区；搜索命中时把当前单元格涂黄
    const selM = this.covered.get(this.key(this.sel.r, this.sel.c));
    const cell = {
      x: X(this.sel.c),
      y: Y(this.sel.r),
      w: colX[(selM?.e.c ?? this.sel.c) + 1] - colX[this.sel.c],
      h: rowY[(selM?.e.r ?? this.sel.r) + 1] - rowY[this.sel.r],
    };
    const b = this.bounds();
    const multi = b.r0 !== b.r1 || b.c0 !== b.c1;
    const selRect = multi
      ? { x: X(b.c0), y: Y(b.r0), w: colX[b.c1 + 1] - colX[b.c0], h: rowY[b.r1 + 1] - rowY[b.r0] }
      : cell;
    if (this.hit) {
      ctx.fillStyle = k['grid-hit'];
      ctx.fillRect(cell.x, cell.y, cell.w, cell.h);
    }

    ctx.font = FONT;
    ctx.textBaseline = 'middle';

    // 合并单元格：擦掉内部网格线，在整块区域里画文字
    for (const m of this.sheet.merges) {
      if (m.e.r < r0 || m.s.r > r1 || m.e.c < c0 || m.s.c > c1) continue;
      const x = X(m.s.c);
      const y = Y(m.s.r);
      const mw = colX[m.e.c + 1] - colX[m.s.c];
      const mh = rowY[m.e.r + 1] - rowY[m.s.r];
      if (!(this.hit && m === selM)) {
        ctx.fillStyle = k['grid-bg'];
        ctx.fillRect(x, y, mw - 1, mh - 1);
      }
      const t = this.text(m.s.r, m.s.c);
      if (t) this.drawText(t, x, y, mw, mh, mw, isNumeric(t) ? 'right' : 'center');
    }

    // 普通单元格（从可视区左侧往前多看几列，让左边溢出过来的长文本也能显示）
    const rows = this.sheet.rows;
    for (let r = r0; r <= r1; r++) {
      const row = rows[r];
      const rh = rowY[r + 1] - rowY[r];
      if (!row || !rh) continue;
      for (let c = Math.max(0, c0 - 3); c <= c1 && c < row.length; c++) {
        const t = row[c];
        if (!t || this.covered.has(this.key(r, c))) continue;
        const cw = colX[c + 1] - colX[c];
        if (!cw) continue;
        const multi = t.includes('\n');
        const align = !multi && isNumeric(t) ? 'right' : 'left';
        let clipW = cw;
        if (align === 'left' && !multi) {
          // 文本超宽且右侧为空时向右溢出，与 Excel 一致
          const tw = ctx.measureText(t).width + PAD * 2;
          for (let cc = c + 1; clipW < tw && cc < this.nCols && cc <= c1 + 1; cc++) {
            if (row[cc] || this.covered.has(this.key(r, cc))) break;
            clipW += colX[cc + 1] - colX[cc];
          }
          if (clipW > cw) {
            ctx.fillStyle = k['grid-bg'];
            ctx.fillRect(X(c) + cw - 1, Y(r), clipW - cw, rh - 1);
          }
        }
        this.drawText(t, X(c), Y(r), cw, rh, clipW, align);
      }
    }
    ctx.restore();

    // 选中框；选区半透明盖在文字上，当前单元格挖空保持原样
    ctx.strokeStyle = k.accent;
    ctx.lineWidth = 2 / this.z;
    ctx.save();
    ctx.beginPath();
    ctx.rect(RW, HEAD_H, w - RW, h - HEAD_H);
    ctx.clip();
    if (multi) {
      ctx.fillStyle = k['grid-range'];
      ctx.beginPath();
      ctx.rect(selRect.x, selRect.y, selRect.w, selRect.h);
      ctx.rect(cell.x, cell.y, cell.w, cell.h);
      ctx.fill('evenodd');
    }
    ctx.strokeRect(selRect.x, selRect.y, selRect.w - 1, selRect.h - 1);
    ctx.restore();

    this.drawHeaders(w, h, r0, r1, c0, c1, X, Y, selRect);
  };

  private drawText(t: string, x: number, y: number, cw: number, rh: number, clipW: number, align: CanvasTextAlign) {
    const ctx = this.ctx;
    ctx.fillStyle = this.colors['grid-text'];
    ctx.textAlign = align;
    const tx = align === 'right' ? x + cw - PAD : align === 'center' ? x + cw / 2 : x + PAD;
    const lines = t.split('\n');
    const needClip = lines.length > 1 || ctx.measureText(t).width > clipW - PAD * 2;
    if (needClip) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(x, y, clipW - 1, rh - 1);
      ctx.clip();
    }
    if (lines.length === 1) {
      ctx.fillText(t, tx, y + rh / 2);
    } else {
      const top = y + Math.max(2, (rh - lines.length * LINE_H) / 2) + LINE_H / 2;
      for (let i = 0; i < lines.length && i * LINE_H < rh; i++) ctx.fillText(lines[i], tx, top + i * LINE_H);
    }
    if (needClip) ctx.restore();
  }

  private drawHeaders(
    w: number, h: number, r0: number, r1: number, c0: number, c1: number,
    X: (c: number) => number, Y: (r: number) => number,
    sel: { x: number; y: number; w: number; h: number },
  ) {
    const { ctx, colors: k, colX, rowY } = this;
    const RW = this.rowNumW;
    ctx.fillStyle = k['grid-head'];
    ctx.fillRect(0, 0, w, HEAD_H);
    ctx.fillRect(0, 0, RW, h);

    // 选中行列在表头上高亮
    ctx.fillStyle = k['accent-soft'];
    ctx.fillRect(Math.max(RW, sel.x), 0, sel.w, HEAD_H);
    ctx.fillRect(0, Math.max(HEAD_H, sel.y), RW, sel.h);

    ctx.font = HEAD_FONT;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.strokeStyle = k['grid-line'];
    ctx.lineWidth = 1 / this.z;

    ctx.save();
    ctx.beginPath();
    ctx.rect(RW, 0, w - RW, HEAD_H);
    ctx.clip();
    ctx.fillStyle = k['grid-head-text'];
    ctx.beginPath();
    for (let c = c0; c <= c1; c++) {
      const cw = colX[c + 1] - colX[c];
      const room = this.f && c < this.dataCols ? FUN_W : 0; // 筛选按钮占掉右侧一块
      if (cw) ctx.fillText(colName(c), X(c) + (cw - room) / 2, HEAD_H / 2);
      const x = Math.round(X(c + 1)) - 0.5;
      ctx.moveTo(x, 0);
      ctx.lineTo(x, HEAD_H);
    }
    ctx.stroke();
    if (this.f) {
      for (let c = c0; c <= c1 && c < this.dataCols; c++) {
        if (colX[c + 1] === colX[c]) continue;
        const on = this.f.cols.has(c);
        const cx = X(c + 1) - FUN_W / 2;
        const cy = HEAD_H / 2;
        if (on) {
          ctx.fillStyle = k['accent-soft'];
          ctx.fillRect(cx - FUN_W / 2 + 1, 3, FUN_W - 2, HEAD_H - 6);
        }
        ctx.fillStyle = on ? k.accent : k['grid-head-text'];
        ctx.beginPath();
        ctx.moveTo(cx - 4.5, cy - 3.5);
        ctx.lineTo(cx + 4.5, cy - 3.5);
        ctx.lineTo(cx + 0.9, cy + 0.5);
        ctx.lineTo(cx + 0.9, cy + 4);
        ctx.lineTo(cx - 0.9, cy + 3);
        ctx.lineTo(cx - 0.9, cy + 0.5);
        ctx.closePath();
        ctx.fill();
      }
    }
    ctx.restore();

    ctx.save();
    ctx.beginPath();
    ctx.rect(0, HEAD_H, RW, h - HEAD_H);
    ctx.clip();
    ctx.fillStyle = k['grid-head-text'];
    ctx.beginPath();
    for (let r = r0; r <= r1; r++) {
      const rh = rowY[r + 1] - rowY[r];
      if (rh) ctx.fillText(String(r + 1), RW / 2, Y(r) + rh / 2);
      const y = Math.round(Y(r + 1)) - 0.5;
      ctx.moveTo(0, y);
      ctx.lineTo(RW, y);
    }
    ctx.stroke();
    ctx.restore();

    ctx.beginPath();
    ctx.moveTo(0, HEAD_H - 0.5);
    ctx.lineTo(w, HEAD_H - 0.5);
    ctx.moveTo(RW - 0.5, 0);
    ctx.lineTo(RW - 0.5, h);
    ctx.stroke();
  }
}

export const renderSheet: Renderer = async (data, { host, shell, name, saved }) => {
  const sheets = await parseInWorker(data, name);
  if (!sheets.length) throw new Error('表格中没有工作表');

  const view = el('div', 'sheet-view');
  const bar = el('div', 'sheet-bar');
  const ref = el('div', 'cell-ref');
  const value = el('div', 'cell-value');
  const filterStat = el('span', 'filter-stat');
  const filterClear = el('button', 'filter-clear', '清除筛选');
  filterStat.hidden = filterClear.hidden = true;
  bar.append(ref, value, filterStat, filterClear);

  // 搜索面板：画布上的内容浏览器自带的查找搜不到，所以自己做
  const panel = el('div', 'search-panel');
  const search = el('input', 'search-input');
  search.placeholder = '在工作表中搜索';
  const findStatus = el('span', 'search-status');
  const closeBtn = el('button', 'search-close');
  closeBtn.title = '关闭 (Esc)';
  closeBtn.append(icon('close'));
  panel.append(icon('search'), search, findStatus, closeBtn);
  panel.hidden = true;

  const gridBox = el('div', 'sheet-grid');
  const scroller = el('div', 'sheet-scroll');
  view.append(bar, gridBox, panel);
  host.append(view);

  const grid = new Grid(scroller, gridBox);
  grid.onSelect = (r, v) => {
    ref.textContent = r;
    value.textContent = v;
    value.title = v;
  };

  grid.onCopy = (rows, cols) => shell.hud(rows === 1 && cols === 1 ? '已复制' : `已复制 ${rows} 行 × ${cols} 列`);
  grid.onFilterChange = (shown, total, active) => {
    filterStat.hidden = !grid.filtering;
    filterStat.textContent = active ? `显示 ${shown.toLocaleString()} / ${total.toLocaleString()} 行` : '点列头右侧的漏斗图标筛选';
    filterClear.hidden = !active;
  };
  filterClear.addEventListener('click', () => {
    grid.clearFilters();
    scroller.focus({ preventScroll: true });
  });
  let closeMenu: (() => void) | null = null;
  grid.onFilterClick = (col, x, y) => {
    closeMenu?.();
    const { values, allowed } = grid.columnValues(col);
    closeMenu = showFilterMenu(gridBox, x, y, {
      title: `${colName(col)} 列`,
      values,
      allowed,
      onApply: (set) => grid.setColumnFilter(col, set),
      onClose: () => {
        closeMenu = null;
        scroller.focus({ preventScroll: true });
      },
    });
  };
  function toggleFilter() {
    closeMenu?.();
    const on = grid.toggleFilter();
    shell.hud(on ? '筛选已开启：点列头右侧的漏斗图标' : '筛选已关闭');
    scroller.focus({ preventScroll: true });
  }

  let current = 0;
  const colsOf = (s: SheetData) => s.rows.reduce((m, row) => Math.max(m, row.length), 0);
  shell.setOutline(
    sheets.map((s, i) => ({ id: String(i), title: s.name, level: 1, meta: `${s.rows.length.toLocaleString()} 行 × ${colsOf(s)} 列` })),
    (id) => show(Number(id)),
  );
  function show(i: number) {
    closeMenu?.();
    current = i;
    shell.setActive(String(i));
    findStatus.textContent = '';
    grid.setSheet(sheets[i]);
    scroller.focus({ preventScroll: true });
  }

  function openSearch() {
    panel.hidden = false;
    search.focus();
    search.select();
  }
  function closeSearch() {
    panel.hidden = true;
    findStatus.textContent = '';
    scroller.focus({ preventScroll: true });
  }
  shell.setActions([
    { icon: 'search', title: '搜索 (⌘/Ctrl + F)', run: openSearch },
    { icon: 'filter', title: '筛选 (⌘/Ctrl + Shift + L)：以当前行为表头', run: toggleFilter },
    { icon: 'copy', title: '复制选区 (⌘/Ctrl + C)', run: () => void grid.copy() },
  ]);
  scroller.addEventListener('keydown', (e) => {
    if ((e.metaKey || e.ctrlKey) && e.shiftKey && e.key.toLowerCase() === 'l') {
      e.preventDefault();
      toggleFilter();
    }
  });
  closeBtn.addEventListener('click', closeSearch);
  search.addEventListener('input', () => (findStatus.textContent = ''));
  search.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') findStatus.textContent = grid.find(search.value, e.shiftKey) ? '' : '无结果';
    else if (e.key === 'Escape') closeSearch();
  });

  // 恢复上次的工作表、缩放、选中格和滚动位置
  show(saved.sheet != null && saved.sheet < sheets.length ? saved.sheet : 0);
  if (saved.zoom) grid.setZoom(saved.zoom);
  if (saved.cell) grid.select(...saved.cell);
  if (saved.sx != null) scroller.scrollLeft = saved.sx;
  if (saved.sy != null) scroller.scrollTop = saved.sy;

  const ro = new ResizeObserver(grid.schedule);
  ro.observe(gridBox);

  return {
    scroller,
    openSearch,
    zoom: { get: () => grid.zoom, set: (z, anchor) => grid.setZoom(z, anchor) },
    state: () => ({ sheet: current, cell: grid.selection, zoom: grid.zoom, sx: scroller.scrollLeft, sy: scroller.scrollTop }),
    onTheme: () => grid.readColors(),
    destroy() {
      closeMenu?.();
      ro.disconnect();
      grid.destroy();
    },
  };
};
