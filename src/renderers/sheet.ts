import { icon } from '../icons';
import { cssVar, el } from '../ui';
import type { CellRange, SheetData, SheetWorkerResponse } from './sheet-types';
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

function parseInWorker(buf: ArrayBuffer, name: string): Promise<SheetData[]> {
  const worker = new Worker(new URL('../workers/sheet.worker.ts', import.meta.url), { type: 'module' });
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
  private sel = { r: 0, c: 0 };
  private hit = false;
  /** 缩放系数：内部一律使用未缩放的逻辑坐标，只在绘制、滚动和鼠标换算时乘除 z */
  private z = 1;

  onSelect: (ref: string, value: string) => void = () => {};

  constructor(private scroller: HTMLElement, gridBox: HTMLElement) {
    gridBox.append(this.canvas, scroller);
    scroller.append(this.spacer);
    scroller.tabIndex = 0;
    scroller.addEventListener('scroll', this.schedule);
    scroller.addEventListener('mousedown', this.onMouseDown);
    scroller.addEventListener('keydown', this.onKeyDown);
    this.readColors();
  }

  readColors() {
    for (const k of ['grid-bg', 'grid-line', 'grid-head', 'grid-head-text', 'grid-text', 'accent', 'accent-soft', 'grid-hit']) {
      this.colors[k] = cssVar(`--${k}`);
    }
    this.schedule();
  }

  setSheet(sheet: SheetData) {
    this.sheet = sheet;
    const dataCols = sheet.rows.reduce((m, row) => Math.max(m, row.length), 0);
    this.nRows = sheet.rows.length + EXTRA_ROWS;
    this.nCols = Math.max(dataCols + EXTRA_COLS, 26);
    this.colX = prefixSums(this.nCols, (c) => sheet.colWidths[c] ?? COL_W);
    this.rowY = prefixSums(this.nRows, (r) => sheet.rowHeights[r] ?? ROW_H);
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
  }

  // ---------- 选择与导航 ----------
  select(r: number, c: number, fromSearch = false) {
    r = Math.max(0, Math.min(r, this.nRows - 1));
    c = Math.max(0, Math.min(c, this.nCols - 1));
    const m = this.covered.get(this.key(r, c));
    if (m) ({ r, c } = m.s);
    this.sel = { r, c };
    this.hit = fromSearch;
    this.onSelect(`${colName(c)}${r + 1}`, this.text(r, c));
    this.ensureVisible(r, c);
    this.schedule();
  }

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

  private onMouseDown = (e: MouseEvent) => {
    const rect = this.scroller.getBoundingClientRect();
    if (e.clientX - rect.left > this.scroller.clientWidth || e.clientY - rect.top > this.scroller.clientHeight) return; // 点在滚动条上
    const px = (e.clientX - rect.left) / this.z;
    const py = (e.clientY - rect.top) / this.z;
    if (px < this.rowNumW || py < HEAD_H) return;
    const c = indexAt(this.colX, px - this.rowNumW + this.scroller.scrollLeft / this.z);
    const r = indexAt(this.rowY, py - HEAD_H + this.scroller.scrollTop / this.z);
    this.select(r, c);
  };

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
    if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'c') {
      navigator.clipboard?.writeText(this.text(r, c));
      e.preventDefault();
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
      if (!row || !row.length) {
        // 整行为空时直接跳过这一行
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

    // 搜索命中高亮
    const selM = this.covered.get(this.key(this.sel.r, this.sel.c));
    const selRect = {
      x: X(this.sel.c),
      y: Y(this.sel.r),
      w: colX[(selM?.e.c ?? this.sel.c) + 1] - colX[this.sel.c],
      h: rowY[(selM?.e.r ?? this.sel.r) + 1] - rowY[this.sel.r],
    };
    if (this.hit) {
      ctx.fillStyle = k['grid-hit'];
      ctx.fillRect(selRect.x, selRect.y, selRect.w, selRect.h);
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

    // 选中框
    ctx.strokeStyle = k.accent;
    ctx.lineWidth = 2 / this.z;
    ctx.save();
    ctx.beginPath();
    ctx.rect(RW, HEAD_H, w - RW, h - HEAD_H);
    ctx.clip();
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
      if (cw) ctx.fillText(colName(c), X(c) + cw / 2, HEAD_H / 2);
      const x = Math.round(X(c + 1)) - 0.5;
      ctx.moveTo(x, 0);
      ctx.lineTo(x, HEAD_H);
    }
    ctx.stroke();
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
  bar.append(ref, value);

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

  let current = 0;
  const colsOf = (s: SheetData) => s.rows.reduce((m, row) => Math.max(m, row.length), 0);
  shell.setOutline(
    sheets.map((s, i) => ({ id: String(i), title: s.name, level: 1, meta: `${s.rows.length.toLocaleString()} 行 × ${colsOf(s)} 列` })),
    (id) => show(Number(id)),
  );
  function show(i: number) {
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
  shell.setActions([{ icon: 'search', title: '搜索 (⌘/Ctrl + F)', run: openSearch }]);
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
      ro.disconnect();
      grid.destroy();
    },
  };
};
