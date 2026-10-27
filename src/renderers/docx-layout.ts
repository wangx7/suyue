// docx-preview 渲染结果的排版修正。阅读视图里正文栏比纸张宽，按 pt 排的表格、图片、
// 右对齐制表位都需要换算成相对正文栏的比例；两种视图共用的修正（行距、单元格边距、表格定宽、编号制表位）也在这里。

import { el } from '../ui';

/** docx-preview 输出的长度都是 “12.34pt”，其他单位返回 NaN */
const pt = (v: string) => (v.endsWith('pt') ? parseFloat(v) : NaN);
const sum = (xs: number[]) => xs.reduce((a, b) => a + b, 0);
const PX_PER_PT = 96 / 72;

/** 阅读视图里正文字号放大到约 16px，最多放大 1.2 倍 */
const READ_TEXT_PX = 16;
const READ_SCALE_MAX = 1.2;

interface Tab {
  style?: string;
  leader?: string;
  position?: string;
}
interface DocStyle {
  id?: string;
  basedOn?: string;
  paragraphProps?: { tabs?: Tab[] };
}

const EDGE = 'docx-edge-tab';

/**
 * 渲染前标记“最后一个制表位是贴着右边距的右对齐制表位”的段落（典型如目录的页码），
 * 阅读视图里由 fitEdgeTabs 改为对齐正文栏右边缘。段落模型会被反复渲染，所以每次都重新设置。
 */
export function markEdgeTabs(doc: any, read: boolean) {
  const props = doc.documentPart?.body?.props;
  const content = pt(props?.pageSize?.width ?? '') - pt(props?.pageMargins?.left ?? '') - pt(props?.pageMargins?.right ?? '');
  const styles: DocStyle[] = doc.stylesPart?.styles ?? [];
  const byId = new Map(styles.map((s) => [s.id, s]));
  const tabsOf = (s: DocStyle | undefined, depth = 0): Tab[] | undefined =>
    !s || depth > 10 ? undefined : (s.paragraphProps?.tabs ?? tabsOf(s.basedOn ? byId.get(s.basedOn) : undefined, depth + 1));

  // 目录常放在排版用的表格里，单元格内按单元格宽度判断
  const walk = (children: any[], width: number) => {
    for (const node of children ?? []) {
      if (node.type === 'table') {
        const cols: number[] = (node.columns ?? []).map((c: { width?: string }) => pt(c.width ?? ''));
        for (const row of node.children ?? []) {
          let i = 0;
          for (const cell of row.children ?? []) {
            const span = cell.span ?? 1;
            walk(cell.children, sum(cols.slice(i, i + span)));
            i += span;
          }
        }
      }
      if (node.type !== 'paragraph' || (node.className && !node.className.startsWith(EDGE))) continue;
      const stops = ((node.tabs as Tab[] | undefined) ?? tabsOf(byId.get(node.styleName)) ?? []).filter((t) => t.style !== 'clear');
      const last = stops.reduce<Tab | undefined>((a, b) => (!a || pt(b.position ?? '') > pt(a.position ?? '') ? b : a), undefined);
      const edge = read && last?.style === 'right' && pt(last.position ?? '') >= width * 0.9;
      node.className = edge ? `${EDGE} docx-leader-${last!.leader ?? 'none'}` : undefined;
    }
  };
  walk(doc.documentPart?.body?.children, content);
}

/**
 * 把标记过的段落拆成“标题 | 前导符 | 页码”三段的 flex 行，页码贴右边缘，窗口缩放时自动跟随。
 * 原来的 tab 元素隐藏后留在原处：docx-preview 渲染后还会异步调整它，脱离文档会让后续制表位都不再处理。
 */
export function fitEdgeTabs(root: HTMLElement) {
  for (const p of root.querySelectorAll<HTMLElement>(`p.${EDGE}`)) {
    const tabs = p.querySelectorAll<HTMLElement>('.docx-tab-stop');
    const tab = tabs[tabs.length - 1];
    if (!tab) continue;
    const indent = parseFloat(getComputedStyle(p).textIndent);
    const range = document.createRange();
    range.setStart(p, 0);
    range.setEndBefore(tab);
    const head = range.extractContents();
    range.setStartAfter(tab);
    range.setEnd(p, p.childNodes.length);
    const tail = range.extractContents();

    const title = el('span', 'docx-tab-title');
    title.append(head);
    if (indent > 0) title.style.paddingInlineStart = `${indent}px`; // flex 容器上 text-indent 不生效
    const fill = el('span', 'docx-tab-fill');
    fill.append(...p.childNodes);
    tab.style.display = 'none';
    const end = el('span', 'docx-tab-end');
    end.append(tail);
    p.replaceChildren(title, fill, end);
    p.classList.add('docx-tab-line');
  }
}

/** 记下所有表格原始的 pt 列宽，并按此算出某个元素所在的“版心”宽度：单元格内容宽，或所在节的版心宽 */
function widthResolver(root: HTMLElement) {
  const grids = new Map<HTMLTableElement, { cols: HTMLTableColElement[]; widths: number[] }>();
  for (const t of root.querySelectorAll<HTMLTableElement>('table')) {
    const cols = [...t.querySelectorAll<HTMLTableColElement>(':scope > colgroup > col')];
    grids.set(t, { cols, widths: cols.map((c) => pt(c.style.width)) });
  }
  const basePt = (node: Element) => {
    const td = node.parentElement?.closest('td');
    const grid = td && grids.get(td.closest('table')!);
    if (td && grid) {
      let i = 0;
      for (const c of td.parentElement!.children) {
        if (c === td) break;
        i += (c as HTMLTableCellElement).colSpan;
      }
      return sum(grid.widths.slice(i, i + td.colSpan)) - (pt(td.style.paddingLeft) || 0) - (pt(td.style.paddingRight) || 0);
    }
    const s = node.closest<HTMLElement>('section.docx');
    return s ? pt(s.style.width) - (pt(s.style.paddingLeft) || 0) - (pt(s.style.paddingRight) || 0) : NaN;
  };
  return { grids, basePt };
}

type Resolver = ReturnType<typeof widthResolver>;

const CJK = /[⺀-鿿豈-﫿＀-￯]/;
/** 很长的无空格串（哈希、URL、Base64），Word 会在行尾直接断开，浏览器却整串挪到下一行、把上一行两端对齐拉得很稀 */
const LONG_TOKEN = /[!-~]{40,}/;
/** 段落里不同字号的文字按字数加权 */
type SizeWeights = Map<number, number>;
const addWeight = (m: SizeWeights, size: number, n: number) => m.set(size, (m.get(size) ?? 0) + n);
const dominant = (m: SizeWeights) => [...m].reduce<[number, number]>((a, b) => (b[1] > a[1] ? b : a), [0, 0])[0];

/** 文档网格的行距（px），没有按行对齐的网格时为 0 */
function gridPitch(doc: any): number {
  const grid = doc.documentPart?.body?.props?.docGrid;
  const pitch = pt(grid?.linePitch ?? '');
  return grid?.type && grid.type !== 'default' && pitch > 0 ? pitch * PX_PER_PT : 0;
}

/**
 * 段落行距、字符单位、编号样式。docx-preview 只给出 Word 的设置（见补丁里的 --docx-line 等），
 * 这里按段落实际文字算出 Word 的效果：
 * - 多倍行距以字体的单倍行高为单位（中文字体约 1.3 倍字号，西文约 1.17 倍），不是 CSS 的字号；
 *   段落对齐文档网格时，单倍行高向上取整到网格行距的整数倍。字号明显更大的文字单独撑高所在行。
 * - “最小值”行距取设定值与单倍行高中的较大者；固定行距不动。
 * - 字符单位缩进按段落的主要字号换算（段落元素本身的字号是文档默认值，文字的字号在 span 上）。
 * - 编号沿用段落首个文字的字号、粗细和字体。
 * - “相同样式的段落间不添加空格”。
 * 返回全文的主要字号（px）。
 */
function fitParagraphs(root: HTMLElement, doc: any): number {
  const pitch = gridPitch(doc);
  const sizes = new Map<Element, number>();
  const sizeOf = (e: Element) => {
    let s = sizes.get(e);
    if (s === undefined) sizes.set(e, (s = parseFloat(getComputedStyle(e).fontSize)));
    return s;
  };
  const styleOf = (p: Element | null) => (p?.tagName === 'P' ? ([...p.classList].find((c) => c.startsWith('docx_')) ?? '') : null);
  const all: SizeWeights = new Map();

  const plans = [];
  for (const p of root.querySelectorAll<HTMLElement>('section.docx p')) {
    const cs = getComputedStyle(p);
    const weights: SizeWeights = new Map();
    const runs = new Map<HTMLElement, number>();
    const long = new Set<HTMLElement>();
    let first: HTMLElement | null = null;
    const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    for (let t = walker.nextNode(); t; t = walker.nextNode()) {
      const text = t.textContent!;
      const n = text.trim().length;
      const el = t.parentElement;
      if (!n || !el) continue;
      const s = sizeOf(el);
      addWeight(weights, s, n);
      runs.set(el, s);
      first ??= el;
      if (LONG_TOKEN.test(text)) long.add(el);
    }
    for (const [s, n] of weights) addWeight(all, s, n);
    const size = dominant(weights) || parseFloat(cs.fontSize);

    const [rule, value] = cs.getPropertyValue('--docx-line').trim().split(/\s+/);
    const factor = CJK.test(p.textContent ?? '') ? 1.3 : 1.17;
    const snap = pitch > 0 && cs.getPropertyValue('--docx-snap').trim() !== '0';
    const single = (s: number) => (snap ? Math.max(1, Math.ceil((s * factor) / pitch - 0.05)) * pitch : s * factor);
    let line = 0;
    const tall: [HTMLElement, number][] = [];
    if (!rule || rule === 'auto') {
      const multiple = Number(value) || 1;
      line = multiple * single(size);
      for (const [el, s] of runs) if (s > size * 1.1 && el !== p) tall.push([el, multiple * single(s)]);
    } else if (rule === 'atLeast') {
      line = Math.max(Number(value) * PX_PER_PT, size * factor);
    }

    const mark = first && [...p.classList].some((c) => c.startsWith('docx-num-')) ? getComputedStyle(first) : null;
    const ctx = cs.getPropertyValue('--docx-ctx').trim() === '1';
    const style = styleOf(p);
    plans.push({
      p,
      size: weights.size ? size : 0,
      line,
      tall,
      long,
      mark: mark && { weight: mark.fontWeight, font: mark.fontFamily },
      noBefore: ctx && styleOf(p.previousElementSibling) === style,
      noAfter: ctx && styleOf(p.nextElementSibling) === style,
    });
  }

  for (const { p, size, line, tall, long, mark, noBefore, noAfter } of plans) {
    if (size) p.style.setProperty('--docx-ch', `${size}px`);
    if (line) p.style.lineHeight = `${line}px`;
    for (const [el, h] of tall) el.style.lineHeight = `${h}px`;
    for (const el of long) el.style.wordBreak = 'break-all';
    if (mark) {
      p.style.setProperty('--docx-mark-weight', mark.weight);
      p.style.setProperty('--docx-mark-font', mark.font);
    }
    if (noBefore) p.style.marginTop = '0';
    if (noAfter) p.style.marginBottom = '0';
  }
  return dominant(all);
}

/**
 * 单元格边距：表格样式里的 tblCellMar 被 docx-preview 写成了 table 的 padding（border-collapse 下不生效），
 * 挪到没有自带边距的单元格上。文档里连默认表格样式都没有时，按 Word 的默认值左右各 5.4pt。
 */
function fitCells(tables: HTMLTableElement[], defaultStyle: boolean) {
  const SIDES = ['top', 'right', 'bottom', 'left'];
  const WORD_DEFAULT = ['0px', '7.2px', '0px', '7.2px'];
  const plans: { el: HTMLElement; pad: string[] }[] = [];
  for (const t of tables) {
    const tcs = getComputedStyle(t);
    let pad = SIDES.map((s) => tcs.getPropertyValue(`padding-${s}`));
    if (pad.every((v) => !parseFloat(v))) {
      if (defaultStyle) continue;
      pad = WORD_DEFAULT;
    }
    plans.push({ el: t, pad: ['0px', '0px', '0px', '0px'] });
    for (const row of t.rows) {
      for (const cell of row.cells) {
        const cs = getComputedStyle(cell);
        const own = SIDES.map((s) => cell.style.getPropertyValue(`padding-${s}`) !== '' || parseFloat(cs.getPropertyValue(`padding-${s}`)) > 0);
        if (own.every(Boolean)) continue;
        plans.push({ el: cell, pad: pad.map((v, i) => (own[i] ? '' : v)) });
      }
    }
  }
  for (const { el, pad } of plans) pad.forEach((v, i) => v && el.style.setProperty(`padding-${SIDES[i]}`, v));
}

/**
 * 表格按 tblGrid 定宽：gridCol 是 Word 保存时算好的实际列宽，按 fixed 布局才不会被长串撑歪。
 * 阅读视图里再把 pt 换算成相对版心（嵌套表格相对父单元格）的百分比。
 */
function fitTables(root: HTMLElement, read: boolean, { grids, basePt }: Resolver) {
  // 先按原始 pt 值算好，再统一写入，嵌套表格才能拿到父表格的原始列宽
  const plans = [...grids].map(([t, { cols, widths }]) => ({ t, cols, widths, total: sum(widths), base: read ? basePt(t) : NaN }));
  for (const { t, cols, widths, total, base } of plans) {
    if (!cols.length || !(total > 0)) continue;
    t.style.tableLayout = 'fixed';
    if (!read) continue;
    // 浮动表格（tblpPr）在流式阅读里按普通表格排，否则 100% 宽再加外边距会超出正文栏
    if (t.style.float) {
      t.style.float = '';
      t.style.marginLeft = t.style.marginRight = '';
    }
    cols.forEach((c, i) => (c.style.width = `${(widths[i] / total) * 100}%`));
    const w = pt(t.style.width) || total;
    if (base > 0 && !t.style.width.endsWith('%')) t.style.width = `${Math.min(100, (w / base) * 100)}%`;
  }
  // 单元格宽度（tcW）只在页面视图使用，阅读视图由列宽决定
  if (read) for (const td of root.querySelectorAll<HTMLElement>('td')) td.style.removeProperty('width');
}

/**
 * 嵌入型图片按“占版心的比例”显示，跟随正文栏缩放；最大不超过原图像素宽度（Word 里本来就放大显示的除外）。
 * 浮动、环绕、旋转、裁剪的图片保持原样。
 */
function fitImages(root: HTMLElement, { basePt }: Resolver, scale: number) {
  const plans = [];
  for (const img of root.querySelectorAll<HTMLImageElement>('section.docx img')) {
    const drawing = img.parentElement;
    if (!drawing || drawing.tagName !== 'DIV' || drawing.style.position !== 'relative' || drawing.style.float) continue;
    if (img.style.transform || img.style.clipPath) continue;
    const inline = drawing.style.display === 'inline-block';
    const block = drawing.style.display === 'block' && drawing.style.width === '100%'; // 上下型环绕，容器已占满一行
    const w = pt(img.style.width);
    const h = pt(img.style.height);
    const base = basePt(img);
    if (!(inline || block) || !(w > 0 && h > 0 && base > 0)) continue;
    plans.push({ img, box: inline ? drawing : img, inline, w, h, pct: Math.min(100, (w / base) * 100) });
  }
  for (const { img, box, inline, w, h, pct } of plans) {
    img.style.width = inline ? '100%' : `${pct}%`;
    img.style.height = 'auto';
    img.style.aspectRatio = `${w} / ${h}`;
    if (inline) {
      box.style.width = `${pct}%`;
      box.style.height = 'auto';
    }
    // 阅读视图整体放大了 scale 倍，上限按放大前换算，避免把图片放大到超过原图
    const cap = () => (box.style.maxWidth = `${Math.max(img.naturalWidth, w * PX_PER_PT) / scale}px`);
    if (img.complete && img.naturalWidth) cap();
    else img.addEventListener('load', cap, { once: true });
  }
}

/**
 * 编号后的制表符：CSS 里的 tab 会折叠成空格。按 Word 的规则算出标记宽度：
 * 悬挂缩进时文字对齐到左缩进，否则跳到标记之后的下一个默认制表位（从正文栏左边起算）。
 */
function fixListMarkers(root: HTMLElement, doc: any) {
  const tabSuffix = new Set<string>(
    (doc.numberingPart?.domNumberings ?? [])
      .filter((n: { suff?: string }) => n.suff === 'tab')
      .map((n: { id: string; level: number }) => `docx-num-${n.id}-${n.level}`),
  );
  if (!tabSuffix.size) return;
  const tabPx = ((pt(doc.settingsPart?.settings?.defaultTabStop ?? '') || 36) * 96) / 72;

  const plans: { p: HTMLElement; width: number }[] = [];
  for (const p of root.querySelectorAll<HTMLElement>('p[class*="docx-num-"]')) {
    if (![...p.classList].some((c) => tabSuffix.has(c))) continue;
    const text = document.createTreeWalker(p, NodeFilter.SHOW_TEXT, (n) =>
      n.textContent!.trim() ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP,
    ).nextNode();
    if (!text) continue;
    const cs = getComputedStyle(p);
    const margin = parseFloat(cs.marginLeft);
    const inset = parseFloat(cs.borderLeftWidth) + parseFloat(cs.paddingLeft);
    const indent = parseFloat(cs.textIndent);
    const left = margin + inset; // 左缩进，相对正文栏
    const start = left + indent; // 首行（标记）起点
    // getBoundingClientRect 含 CSS zoom，用渲染宽度和计算宽度之比换算回 CSS px
    const box = p.getBoundingClientRect();
    const scale = box.width / (parseFloat(cs.width) + inset + parseFloat(cs.borderRightWidth) + parseFloat(cs.paddingRight));
    const range = document.createRange();
    range.setStart(text, text.textContent!.search(/\S/));
    // 标记是 inline-block，末尾的 tab 已被去掉，第一个字的位置就是标记的结束处
    const end = margin + (range.getBoundingClientRect().left - box.left) / scale;
    if (!(scale > 0) || !Number.isFinite(end)) continue;
    const stop = indent < 0 && end <= left ? left : (Math.floor(end / tabPx) + 1) * tabPx;
    plans.push({ p, width: stop - start });
  }
  for (const { p, width } of plans) p.style.setProperty('--docx-marker', `${width}px`);
}

/** 阅读视图正文两端对齐。表格内、含制表位的段落不动（制表位按左对齐算的位置） */
function justify(root: HTMLElement) {
  const targets = [...root.querySelectorAll<HTMLElement>('section.docx p')].filter(
    (p) => !p.closest('td') && !p.querySelector('.docx-tab-stop') && /^(start|left)$/.test(getComputedStyle(p).textAlign),
  );
  for (const p of targets) p.style.textAlign = 'justify';
}

/** 渲染后、深色重映射之前调用 */
export function fitLayout(root: HTMLElement, doc: any, read: boolean) {
  const textPx = fitParagraphs(root, doc);
  // 阅读视图按正文字号整体放大（zoom 在节上，表格、图片的百分比宽度和缩进一起缩放）
  const scale = read && textPx > 0 ? Math.min(READ_SCALE_MAX, Math.max(1, READ_TEXT_PX / textPx)) : 1;
  if (read) for (const s of root.querySelectorAll<HTMLElement>('section.docx')) s.style.zoom = String(scale);

  const resolver = widthResolver(root);
  fitTables(root, read, resolver);
  const defaultTableStyle = (doc.stylesPart?.styles ?? []).some((s: { target?: string; isDefault?: boolean }) => s.target === 'table' && s.isDefault);
  fitCells([...resolver.grids.keys()], defaultTableStyle);
  if (read) {
    fitImages(root, resolver, scale);
    justify(root);
  }
  fixListMarkers(root, doc);
}
