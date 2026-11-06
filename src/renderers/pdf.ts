import { getDocument, PDFWorker, TextLayer, type PDFDocumentProxy, type RenderTask } from 'pdfjs-dist';
import workerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?worker&url';
import { icon } from '../icons';
import type { OutlineItem } from '../shell';
import { el } from '../ui';
import type { Renderer } from './types';

// PDF：用 pdf.js 解析，每页一个 Canvas 加一层透明文字（可选中复制），连续竖排滚动。
// 页面只在进入视口前后两屏时才绘制，滚远了释放，几百页的文档也不会占满内存。

const MAX_WIDTH = 1000;
const SIDE_PAD = 24;
/** 与 styles.css 中 .pdf-pages 的上内边距和页间距保持一致 */
const TOP_PAD = 56;
const GAP = 12;
const MAX_PIXELS = 16e6;

async function createWorker(): Promise<Worker> {
  try {
    return new Worker(workerUrl, { type: 'module' });
  } catch {
    // VS Code 的 webview 里脚本和页面不同源，不能直接建 Worker，取回来用 blob 地址建
    const code = await (await fetch(workerUrl)).text();
    return new Worker(URL.createObjectURL(new Blob([code], { type: 'text/javascript' })), { type: 'module' });
  }
}

// 中文 PDF 常引用预定义的 CMap，没内嵌字体时要用标准字体；这些数据随包打进来，按需取
const dataFiles = import.meta.glob(
  ['/node_modules/pdfjs-dist/cmaps/*.bcmap', '/node_modules/pdfjs-dist/standard_fonts/*.{pfb,ttf}', '/node_modules/pdfjs-dist/wasm/*.wasm'],
  { query: '?url&no-inline', import: 'default' },
) as Record<string, () => Promise<string>>;
const DATA_DIR = { cMapUrl: 'cmaps', standardFontDataUrl: 'standard_fonts', wasmUrl: 'wasm' } as const;

class BundledData {
  async fetch({ kind, filename }: { kind: keyof typeof DATA_DIR; filename: string }) {
    const load = dataFiles[`/node_modules/pdfjs-dist/${DATA_DIR[kind]}/${filename}`];
    if (!load) throw new Error(`缺少 PDF 数据文件：${filename}`);
    return new Uint8Array(await (await fetch(await load())).arrayBuffer());
  }
}

interface PageView {
  n: number;
  el: HTMLElement;
  /** 缩放 1 时的页面尺寸（已含页面旋转） */
  w: number;
  h: number;
  canvas?: HTMLCanvasElement;
  layer?: TextLayer;
  layerEl?: HTMLElement;
  task?: RenderTask;
  token: number;
  /** 正在绘制或已绘制的缩放 */
  pending: number;
  renderedScale: number;
  ready?: Promise<void>;
}

interface PageText {
  strs: string[];
  /** 每个文字块在整页文本中的起点 */
  starts: number[];
  full: string;
}

type Dest = { page: number; y: number | null };

export const renderPdf: Renderer = async (data, { host, shell, saved }) => {
  const worker = PDFWorker.create({ port: await createWorker() });
  const task = getDocument({
    data: new Uint8Array(data),
    worker,
    BinaryDataFactory: BundledData,
  });
  let doc: PDFDocumentProxy;
  try {
    doc = await task.promise;
  } catch (err) {
    void task.destroy();
    worker.destroy();
    if ((err as Error).name === 'PasswordException') throw new Error('这个 PDF 设置了密码，暂不支持打开');
    throw err;
  }
  const total = doc.numPages;
  if (!total) throw new Error('这个 PDF 没有页面');

  const scroller = el('div', 'doc-scroll paper');
  const pagesBox = el('div', 'pdf-pages');
  scroller.append(pagesBox);
  host.append(scroller);

  // ---------- 页面与布局 ----------
  const first = (await doc.getPage(1)).getViewport({ scale: 1 });
  const views: PageView[] = [];
  for (let n = 1; n <= total; n++) {
    const page = el('div', 'pdf-page');
    page.dataset.page = String(n);
    pagesBox.append(page);
    views.push({ n, el: page, w: first.width, h: first.height, token: 0, pending: 0, renderedScale: 0 });
  }

  let destroyed = false;
  let zoomLevel = saved.zoom ?? 1;
  let base = 1;
  let scale = 1;
  const tops = new Float64Array(total + 1);

  const fit = () => Math.min(MAX_WIDTH, Math.max(160, scroller.clientWidth - SIDE_PAD * 2)) / first.width;
  function layout() {
    scale = base * zoomLevel;
    let y = TOP_PAD;
    for (const v of views) {
      v.el.style.width = `${v.w * scale}px`;
      v.el.style.height = `${v.h * scale}px`;
      tops[v.n - 1] = y;
      y += v.h * scale + GAP;
    }
    tops[total] = y;
    pagesBox.style.setProperty('--total-scale-factor', String(scale));
  }
  function measureTops() {
    let y = TOP_PAD;
    for (const v of views) {
      tops[v.n - 1] = y;
      y += v.h * scale + GAP;
    }
    tops[total] = y;
  }
  /** 满足 tops[i] <= y 的最大页下标 */
  function pageAt(y: number): number {
    let lo = 0;
    let hi = total - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (tops[mid] <= y) lo = mid;
      else hi = mid - 1;
    }
    return lo;
  }
  const anchor = () => {
    const i = pageAt(scroller.scrollTop);
    return { i, off: (scroller.scrollTop - tops[i]) / scale };
  };
  const restore = (a: { i: number; off: number }) => {
    scroller.scrollTop = tops[a.i] + a.off * scale;
  };

  base = fit();
  layout();

  // 后台读取各页的真实尺寸（页面大小不一的文档才会用到），在当前位置之上的页变化时修正滚动，避免内容跳动
  void (async () => {
    for (let n = 2; n <= total && !destroyed; n++) {
      const v = views[n - 1];
      const vp = (await doc.getPage(n)).getViewport({ scale: 1 });
      if (vp.width !== v.w || vp.height !== v.h) {
        const a = anchor();
        v.w = vp.width;
        v.h = vp.height;
        v.el.style.width = `${v.w * scale}px`;
        v.el.style.height = `${v.h * scale}px`;
        measureTops();
        if (n - 1 < a.i) restore(a);
      }
      if (n % 25 === 0) await new Promise((r) => setTimeout(r));
    }
  })();

  // ---------- 绘制 ----------
  const texts = new Map<number, Promise<PageText>>();
  let query = '';
  const hitCache = new Map<number, number[][]>();
  let cur = { page: 0, k: -1 };

  function pageText(n: number): Promise<PageText> {
    let p = texts.get(n);
    if (!p) {
      p = doc.getPage(n).then(async (page) => {
        const strs: string[] = [];
        const starts: number[] = [];
        let full = '';
        for (const item of (await page.getTextContent()).items) {
          if (!('str' in item)) continue;
          strs.push(item.str);
          starts.push(full.length);
          full += item.str + (item.hasEOL ? '\n' : '');
        }
        return { strs, starts, full: full.toLowerCase() };
      });
      texts.set(n, p);
    }
    return p;
  }

  /** 每处命中覆盖到的文字块下标 */
  async function hitsOf(n: number): Promise<number[][]> {
    let hits = hitCache.get(n);
    if (hits) return hits;
    const { strs, starts, full } = await pageText(n);
    hits = [];
    let from = 0;
    for (let at = full.indexOf(query, from); at >= 0; at = full.indexOf(query, from)) {
      const end = at + query.length;
      const idx: number[] = [];
      for (let i = 0; i < strs.length; i++) {
        if (starts[i] >= end) break;
        if (starts[i] + strs[i].length > at && strs[i]) idx.push(i);
      }
      if (idx.length) hits.push(idx);
      from = end;
    }
    hitCache.set(n, hits);
    return hits;
  }

  function paintHits(v: PageView) {
    const divs = v.layer?.textDivs;
    if (!divs) return;
    for (const d of divs) d.classList.remove('pdf-hit', 'pdf-hit-cur');
    const hits = query ? hitCache.get(v.n) : undefined;
    hits?.forEach((idx, k) => {
      for (const i of idx) divs[i]?.classList.add(cur.page === v.n && cur.k === k ? 'pdf-hit-cur' : 'pdf-hit');
    });
  }

  function free(v: PageView) {
    v.token++;
    v.task?.cancel();
    v.task = undefined;
    v.layer?.cancel();
    v.layer = undefined;
    v.layerEl?.remove();
    v.layerEl = undefined;
    if (v.canvas) {
      v.canvas.width = 0; // 立即释放位图内存
      v.canvas.remove();
      v.canvas = undefined;
    }
    v.ready = undefined;
    v.pending = 0;
    v.renderedScale = 0;
  }

  function render(v: PageView): Promise<void> {
    if (v.ready && v.pending === scale) return v.ready;
    const token = ++v.token;
    v.task?.cancel();
    const target = scale;
    v.pending = target;
    v.ready = (async () => {
      try {
        const page = await doc.getPage(v.n);
        if (token !== v.token) return;
        let out = target * Math.min(window.devicePixelRatio || 1, 3);
        let vp = page.getViewport({ scale: out });
        if (vp.width * vp.height > MAX_PIXELS) {
          out *= Math.sqrt(MAX_PIXELS / (vp.width * vp.height));
          vp = page.getViewport({ scale: out });
        }
        const canvas = el('canvas', 'pdf-canvas');
        canvas.width = Math.floor(vp.width);
        canvas.height = Math.floor(vp.height);
        v.task = page.render({ canvas, viewport: vp });
        await v.task.promise;
        if (token !== v.token) return;
        v.canvas?.remove();
        v.canvas = canvas;
        v.renderedScale = target;
        v.el.prepend(canvas);

        if (!v.layer) {
          const layerEl = el('div', 'textLayer');
          v.el.append(layerEl);
          v.layerEl = layerEl;
          v.layer = new TextLayer({ textContentSource: page.streamTextContent(), container: layerEl, viewport: page.getViewport({ scale: target }) });
          await v.layer.render();
          if (token !== v.token) return;
          if (query) await hitsOf(v.n);
          paintHits(v);
        }
      } catch (err) {
        if ((err as Error).name !== 'RenderingCancelledException' && token === v.token) console.error(err);
      }
    })();
    return v.ready;
  }

  const io = new IntersectionObserver(
    (entries) => {
      for (const e of entries) {
        const v = views[Number((e.target as HTMLElement).dataset.page) - 1];
        if (e.isIntersecting) void render(v);
        else free(v);
      }
    },
    { root: scroller, rootMargin: '200% 0px' },
  );
  views.forEach((v) => io.observe(v.el));

  // ---------- 跳转 ----------
  function scrollToPage(n: number, offset = 0) {
    const i = Math.min(Math.max(n, 1), total) - 1;
    scroller.scrollTop = Math.max(0, tops[i] + offset * scale - 12);
  }

  async function resolve(dest: unknown): Promise<Dest | null> {
    try {
      const d = typeof dest === 'string' ? await doc.getDestination(dest) : dest;
      if (!Array.isArray(d)) return null;
      const ref = d[0];
      const page = typeof ref === 'object' && ref !== null ? await doc.getPageIndex(ref) : Number(ref);
      const mode = d[1]?.name;
      const y = mode === 'XYZ' ? d[3] : mode === 'FitH' || mode === 'FitBH' ? d[2] : null;
      return { page, y: typeof y === 'number' ? y : null };
    } catch {
      return null;
    }
  }
  async function goto(d: Dest) {
    let offset = 0;
    if (d.y != null) offset = (await doc.getPage(d.page + 1)).getViewport({ scale: 1 }).convertToViewportPoint(0, d.y)[1];
    scrollToPage(d.page + 1, offset);
  }

  // ---------- 目录 ----------
  type OutlineNode = { title: string; dest: unknown; items?: OutlineNode[] };
  const nodes: { title: string; level: number; dest: unknown }[] = [];
  const walk = (list: OutlineNode[], level: number) => {
    for (const n of list) {
      if (nodes.length >= 3000) return;
      nodes.push({ title: n.title.trim() || '（无标题）', level, dest: n.dest });
      if (n.items?.length) walk(n.items, level + 1);
    }
  };
  walk(((await doc.getOutline().catch(() => null)) ?? []) as OutlineNode[], 1);

  const dests: (Dest | null)[] = [];
  let activeOutline = (page: number) => shell.setActive(String(page));
  if (nodes.length) {
    const items: OutlineItem[] = nodes.map((n, i) => ({ id: String(i), title: n.title, level: n.level }));
    shell.setOutline(items, async (id) => {
      const i = Number(id);
      dests[i] ??= await resolve(nodes[i].dest);
      const d = dests[i];
      if (d) void goto(d);
    });
    void (async () => {
      for (let i = 0; i < nodes.length && !destroyed; i++) dests[i] = await resolve(nodes[i].dest);
    })();
    activeOutline = (page) => {
      let hit = -1;
      for (let i = 0; i < dests.length; i++) if (dests[i] && dests[i]!.page + 1 <= page) hit = i;
      if (hit >= 0) shell.setActive(String(hit));
    };
  } else {
    const items: OutlineItem[] = views.map((v) => ({ id: String(v.n), title: `第 ${v.n} 页`, level: 1 }));
    shell.setOutline(items, (id) => scrollToPage(Number(id)));
  }

  // ---------- 页码 ----------
  const pager = el('div', 'pdf-pager');
  const pageInput = el('input', 'pdf-pager-input');
  pageInput.inputMode = 'numeric';
  pageInput.title = '跳转到页码';
  pager.append(pageInput, el('span', '', `/ ${total}`));
  host.append(pager);

  let curPage = 1;
  let raf = 0;
  const onScroll = () => {
    if (raf) return;
    raf = requestAnimationFrame(() => {
      raf = 0;
      curPage = pageAt(scroller.scrollTop + scroller.clientHeight / 3) + 1;
      if (document.activeElement !== pageInput) pageInput.value = String(curPage);
      activeOutline(curPage);
    });
  };
  scroller.addEventListener('scroll', onScroll, { passive: true });
  pageInput.addEventListener('focus', () => pageInput.select());
  pageInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      const n = parseInt(pageInput.value, 10);
      if (n >= 1) scrollToPage(n);
      pageInput.blur();
    } else if (e.key === 'Escape') pageInput.blur();
  });
  pageInput.addEventListener('blur', () => (pageInput.value = String(curPage)));
  pageInput.value = '1';

  // ---------- 搜索 ----------
  const panel = el('div', 'search-panel');
  const search = el('input', 'search-input');
  search.placeholder = '在 PDF 中搜索';
  const status = el('span', 'search-status');
  const closeBtn = el('button', 'search-close');
  closeBtn.title = '关闭 (Esc)';
  closeBtn.append(icon('close'));
  panel.append(icon('search'), search, status, closeBtn);
  panel.hidden = true;
  host.append(panel);

  function clearHits() {
    query = '';
    hitCache.clear();
    cur = { page: 0, k: -1 };
    views.forEach(paintHits);
  }
  function openSearch() {
    panel.hidden = false;
    search.focus();
    search.select();
  }
  function closeSearch() {
    panel.hidden = true;
    status.textContent = '';
    clearHits();
    scroller.focus({ preventScroll: true });
  }

  let searching = 0;
  async function find(text: string, backward: boolean) {
    const q = text.trim().toLowerCase();
    if (!q) return clearHits();
    const run = ++searching;
    if (q !== query) {
      query = q;
      hitCache.clear();
      cur = { page: curPage, k: backward ? Infinity : -1 };
    }
    status.textContent = '搜索中…';
    let p = cur.page || curPage;
    let k = cur.k;
    for (let n = 0; n <= total; n++) {
      const hits = await hitsOf(p);
      if (run !== searching) return;
      k = backward ? Math.min(k, hits.length) - 1 : k + 1;
      if (k >= 0 && k < hits.length) {
        cur = { page: p, k };
        status.textContent = `第 ${p} 页`;
        scrollToPage(p);
        const v = views[p - 1];
        await render(v);
        if (run !== searching) return;
        paintHits(v);
        v.layer?.textDivs[hits[k][0]]?.scrollIntoView({ block: 'center', inline: 'nearest' });
        return;
      }
      p = backward ? (p === 1 ? total : p - 1) : p === total ? 1 : p + 1;
      k = backward ? Infinity : -1;
    }
    cur = { page: 0, k: -1 };
    views.forEach(paintHits);
    status.textContent = '无结果';
  }
  closeBtn.addEventListener('click', closeSearch);
  search.addEventListener('input', () => {
    status.textContent = '';
    if (!search.value.trim()) clearHits();
  });
  search.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') void find(search.value, e.shiftKey);
    else if (e.key === 'Escape') closeSearch();
  });
  shell.setActions([{ icon: 'search', title: '搜索 (⌘/Ctrl + F)', run: openSearch }]);

  // ---------- 缩放与窗口变化 ----------
  let rerender = 0;
  // 先拉伸旧位图保持连续，稍后在新缩放下重绘已有位图的页
  function scheduleRerender() {
    clearTimeout(rerender);
    rerender = window.setTimeout(() => {
      for (const v of views) if (v.canvas && v.pending !== scale) void render(v);
    }, 150);
  }

  let lastWidth = scroller.clientWidth;
  const ro = new ResizeObserver(() => {
    if (scroller.clientWidth === lastWidth) return;
    lastWidth = scroller.clientWidth;
    const a = anchor();
    base = fit();
    layout();
    restore(a);
    scheduleRerender();
  });
  ro.observe(scroller);

  // 恢复上次的位置
  if (saved.page) scrollToPage(saved.page, saved.pageY ?? 0);
  onScroll();

  return {
    scroller,
    openSearch,
    zoom: {
      get: () => zoomLevel,
      set(z, pt) {
        const ratio = z / zoomLevel;
        const px = scroller.scrollLeft + pt.x;
        const py = scroller.scrollTop + pt.y;
        zoomLevel = z;
        layout();
        scroller.scrollLeft = px * ratio - pt.x;
        scroller.scrollTop = py * ratio - pt.y;
        scheduleRerender();
      },
    },
    state: () => {
      const a = anchor();
      return { zoom: zoomLevel, page: a.i + 1, pageY: Math.round(a.off) };
    },
    destroy() {
      destroyed = true;
      clearTimeout(rerender);
      cancelAnimationFrame(raf);
      ro.disconnect();
      io.disconnect();
      views.forEach(free);
      void task.destroy();
      worker.destroy();
    },
  };
};
