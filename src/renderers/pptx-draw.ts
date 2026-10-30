import { el } from '../ui';
import { expand } from './docx-fonts';
import {
  colorOf, EMU, lineOf, loadPart, paintFrom, paintOf, refStyle,
  type ColorCtx, type Deck, type Line, type Paint, type Part, type SlideRef,
} from './pptx-parse';
import { dig, find, kid, kids, num, nsAttr } from './zip';

// 把一页幻灯片画成 DOM：每个形状一个绝对定位的盒子，几何形状用 CSS（复杂的用 SVG），文字用普通段落。
// 图片只记下包内路径（data-src / data-bg），等幻灯片滚到附近再由 pptx.ts 加载。

interface DrawCtx extends ColorCtx {
  deck: Deck;
  slide: SlideRef;
  /** 页码，从 1 开始 */
  num: number;
}

interface Rect {
  x: number;
  y: number;
  w: number;
  h: number;
}
/** EMU 坐标 → 当前容器内的像素坐标（组合里会缩放、平移） */
type Mapper = (x: number, y: number, w: number, h: number) => Rect;
const rootMap: Mapper = (x, y, w, h) => ({ x: x / EMU, y: y / EMU, w: w / EMU, h: h / EMU });

interface Xfrm extends Rect {
  rot: number;
  flipH: boolean;
  flipV: boolean;
  ch?: Rect;
}

const SVG = 'http://www.w3.org/2000/svg';
const svg = <K extends keyof SVGElementTagNameMap>(tag: K) => document.createElementNS(SVG, tag);

function xfrmOf(node: Element | null): Xfrm | null {
  const off = kid(node, 'off');
  const ext = kid(node, 'ext');
  if (!off || !ext) return null;
  const chOff = kid(node, 'chOff');
  const chExt = kid(node, 'chExt');
  return {
    x: num(off, 'x', 0),
    y: num(off, 'y', 0),
    w: num(ext, 'cx', 0),
    h: num(ext, 'cy', 0),
    rot: num(node, 'rot', 0) / 60000,
    flipH: node!.getAttribute('flipH') === '1',
    flipV: node!.getAttribute('flipV') === '1',
    ch: chOff && chExt ? { x: num(chOff, 'x', 0), y: num(chOff, 'y', 0), w: num(chExt, 'cx', 0), h: num(chExt, 'cy', 0) } : undefined,
  };
}

function place(box: HTMLElement, r: Rect, rot = 0) {
  Object.assign(box.style, { left: `${r.x}px`, top: `${r.y}px`, width: `${r.w}px`, height: `${r.h}px` });
  if (rot) box.style.transform = `rotate(${rot}deg)`;
}

// ---------- 占位符 ----------

interface Ph {
  type: string;
  idx: string | null;
}

function phOf(shape: Element): Ph | null {
  const nv = kids(shape).find((k) => k.localName.startsWith('nv'));
  const ph = dig(nv, 'nvPr', 'ph');
  return ph ? { type: ph.getAttribute('type') ?? 'obj', idx: ph.getAttribute('idx') } : null;
}

const isTitle = (t: string) => t === 'title' || t === 'ctrTitle';
/** 母版上只有 title、body、dt、ftr、sldNum 几种占位符 */
const masterType = (t: string) => (isTitle(t) ? 'title' : ['dt', 'ftr', 'sldNum'].includes(t) ? t : 'body');

const spTree = (part: Part | null) => dig(part?.root, 'cSld', 'spTree');

/** 在版式（按 idx，再按类型）或母版（按类型）里找对应的占位符 */
function findPh(part: Part | null, ph: Ph, master: boolean): Element | null {
  const shapes = kids(spTree(part)).map((s) => ({ s, p: phOf(s) })).filter((x): x is { s: Element; p: Ph } => !!x.p);
  if (master) return shapes.find((x) => masterType(x.p.type) === masterType(ph.type))?.s ?? null;
  if (ph.idx != null) {
    const byIdx = shapes.find((x) => x.p.idx === ph.idx);
    if (byIdx) return byIdx.s;
  }
  return shapes.find((x) => x.p.type === ph.type || (isTitle(x.p.type) && isTitle(ph.type)))?.s ?? null;
}

/** 形状自身 + 版式占位符 + 母版占位符（只有幻灯片上的占位符才继承） */
function chainOf(node: Element, part: Part, d: DrawCtx): { chain: Element[]; ph: Ph | null } {
  const ph = phOf(node);
  if (!ph || part !== d.slide.part) return { chain: [node], ph };
  const chain = [node, findPh(d.slide.layout, ph, false), findPh(d.slide.master, ph, true)].filter((n): n is Element => !!n);
  return { chain, ph };
}

const firstOf = <T>(list: Element[], pick: (n: Element) => T | null | undefined): T | undefined => {
  for (const n of list) {
    const v = pick(n);
    if (v != null) return v;
  }
  return undefined;
};

// ---------- 几何形状 ----------

/** 预设形状的调整值（a:avLst 里的 gd），按 min(宽, 高) 的十万分之几 */
function adj(geom: Element | null, name: string, def: number): number {
  for (const g of kids(kid(geom, 'avLst'), 'gd')) {
    if (g.getAttribute('name') === name) {
      const m = /^val\s+(-?\d+)/.exec(g.getAttribute('fmla') ?? '');
      if (m) return Number(m[1]);
    }
  }
  return def;
}

/** 能用 clip-path 多边形表示的预设形状；返回 [x, y] 像素点 */
function polygon(prst: string, w: number, h: number, geom: Element | null): [number, number][] | null {
  const m = Math.min(w, h);
  const a = (name: string, def: number) => (adj(geom, name, def) / 100000) * m;
  switch (prst) {
    case 'triangle': {
      const x = (adj(geom, 'adj', 50000) / 100000) * w;
      return [[x, 0], [w, h], [0, h]];
    }
    case 'rtTriangle':
      return [[0, 0], [w, h], [0, h]];
    case 'diamond':
      return [[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]];
    case 'parallelogram': {
      const x = a('adj', 25000);
      return [[x, 0], [w, 0], [w - x, h], [0, h]];
    }
    case 'trapezoid': {
      const x = a('adj', 25000);
      return [[x, 0], [w - x, 0], [w, h], [0, h]];
    }
    case 'pentagon':
      return [[w / 2, 0], [w, h * 0.38], [w * 0.81, h], [w * 0.19, h], [0, h * 0.38]];
    case 'hexagon': {
      const x = a('adj', 25000);
      return [[x, 0], [w - x, 0], [w, h / 2], [w - x, h], [x, h], [0, h / 2]];
    }
    case 'octagon': {
      const x = a('adj', 29289);
      return [[x, 0], [w - x, 0], [w, x], [w, h - x], [w - x, h], [x, h], [0, h - x], [0, x]];
    }
    case 'homePlate': {
      const x = a('adj', 50000);
      return [[0, 0], [w - x, 0], [w, h / 2], [w - x, h], [0, h]];
    }
    case 'chevron': {
      const x = a('adj', 50000);
      return [[0, 0], [w - x, 0], [w, h / 2], [w - x, h], [0, h], [x, h / 2]];
    }
    case 'rightArrow':
    case 'leftArrow': {
      const t = (h * (1 - adj(geom, 'adj1', 50000) / 100000)) / 2;
      const x = w - a('adj2', 50000);
      const pts: [number, number][] = [[0, t], [x, t], [x, 0], [w, h / 2], [x, h], [x, h - t], [0, h - t]];
      return prst === 'leftArrow' ? pts.map(([px, py]) => [w - px, py]) : pts;
    }
    case 'upArrow':
    case 'downArrow': {
      const t = (w * (1 - adj(geom, 'adj1', 50000) / 100000)) / 2;
      const y = a('adj2', 50000);
      const pts: [number, number][] = [[t, h], [t, y], [0, y], [w / 2, 0], [w, y], [w - t, y], [w - t, h]];
      return prst === 'downArrow' ? pts.map(([px, py]) => [px, h - py]) : pts;
    }
  }
  return null;
}

const DASH: Record<string, number[]> = {
  dash: [4, 3], sysDash: [3, 1], lgDash: [8, 3], dashDot: [4, 3, 1, 3], lgDashDot: [8, 3, 1, 3], sysDot: [1, 1], dot: [1, 1], sysDashDot: [3, 1, 1, 1],
};

function borderCss(line: Line): string {
  const style = line.dash === 'solid' ? 'solid' : /dot/i.test(line.dash) && !/dash/i.test(line.dash) ? 'dotted' : 'dashed';
  return `${line.width}px ${style} ${line.color}`;
}

let markerId = 0;

function strokeSvg(node: SVGElement, line: Line, root: SVGSVGElement) {
  node.setAttribute('stroke', line.color);
  node.setAttribute('stroke-width', String(line.width));
  const dash = DASH[line.dash];
  if (dash) node.setAttribute('stroke-dasharray', dash.map((n) => n * line.width).join(' '));
  if (!line.head && !line.tail) return;
  // 箭头：同一个标记，起点处反向
  const id = `pptx-arrow-${++markerId}`;
  const defs = svg('defs');
  const marker = svg('marker');
  Object.entries({ id, viewBox: '0 0 10 10', refX: '8', refY: '5', markerWidth: '4', markerHeight: '4', orient: 'auto-start-reverse' }).forEach(([k, v]) => marker.setAttribute(k, v));
  const tip = svg('path');
  tip.setAttribute('d', 'M0 0L10 5L0 10z');
  tip.setAttribute('fill', line.color);
  marker.append(tip);
  defs.append(marker);
  root.prepend(defs);
  if (line.tail) node.setAttribute('marker-end', `url(#${id})`);
  if (line.head) node.setAttribute('marker-start', `url(#${id})`);
}

/** 自由形状（custGeom）转成 SVG path */
function customPath(geom: Element, w: number, h: number): string {
  let out = '';
  for (const path of kids(kid(geom, 'pathLst'), 'path')) {
    const sx = num(path, 'w', 0) ? w / num(path, 'w', 1) : 1 / EMU;
    const sy = num(path, 'h', 0) ? h / num(path, 'h', 1) : 1 / EMU;
    const pt = (p: Element | null) => [num(p, 'x', 0) * sx, num(p, 'y', 0) * sy];
    let cx = 0;
    let cy = 0;
    for (const cmd of kids(path)) {
      const pts = kids(cmd, 'pt').map(pt);
      switch (cmd.localName) {
        case 'moveTo':
        case 'lnTo':
          [cx, cy] = pts[0] ?? [cx, cy];
          out += `${cmd.localName === 'moveTo' ? 'M' : 'L'}${cx} ${cy}`;
          break;
        case 'cubicBezTo':
        case 'quadBezTo':
          out += (cmd.localName === 'cubicBezTo' ? 'C' : 'Q') + pts.map((p) => p.join(' ')).join(' ');
          [cx, cy] = pts[pts.length - 1] ?? [cx, cy];
          break;
        case 'arcTo': {
          const rx = num(cmd, 'wR', 0) * sx;
          const ry = num(cmd, 'hR', 0) * sy;
          const st = (num(cmd, 'stAng', 0) / 60000) * (Math.PI / 180);
          const sw = (num(cmd, 'swAng', 0) / 60000) * (Math.PI / 180);
          const ox = cx - rx * Math.cos(st);
          const oy = cy - ry * Math.sin(st);
          cx = ox + rx * Math.cos(st + sw);
          cy = oy + ry * Math.sin(st + sw);
          out += `A${rx} ${ry} 0 ${Math.abs(sw) > Math.PI ? 1 : 0} ${sw > 0 ? 1 : 0} ${cx} ${cy}`;
          break;
        }
        case 'close':
          out += 'Z';
          break;
      }
    }
  }
  return out;
}

/** 连接线的折线路径 */
function connectorPath(prst: string, w: number, h: number, geom: Element | null): string | null {
  if (prst === 'line' || prst === 'straightConnector1') return `M0 0L${w} ${h}`;
  if (prst === 'bentConnector2') return `M0 0L${w} 0L${w} ${h}`;
  if (prst.startsWith('bentConnector')) {
    const x = (adj(geom, 'adj1', 50000) / 100000) * w;
    return `M0 0L${x} 0L${x} ${h}L${w} ${h}`;
  }
  if (prst.startsWith('curvedConnector')) return `M0 0C${w / 2} 0 ${w / 2} ${h} ${w} ${h}`;
  return null;
}

/** 填充和轮廓画在形状盒子的底层；翻转只作用于这一层，文字不跟着镜像 */
function drawGeom(box: HTMLElement, spPr: Element | null, r: Rect, xf: Xfrm, fill: Paint | undefined, line: Line | null | undefined) {
  const prstGeom = kid(spPr, 'prstGeom');
  const custGeom = kid(spPr, 'custGeom');
  const prst = prstGeom?.getAttribute('prst') ?? 'rect';
  const flip = `${xf.flipH ? 'scaleX(-1) ' : ''}${xf.flipV ? 'scaleY(-1)' : ''}`.trim();

  const path = custGeom ? customPath(custGeom, r.w, r.h) : connectorPath(prst, r.w, r.h, prstGeom);
  if (path) {
    const root = svg('svg');
    root.classList.add('pptx-svg');
    root.setAttribute('width', String(Math.max(r.w, 1)));
    root.setAttribute('height', String(Math.max(r.h, 1)));
    if (flip) root.style.transform = flip;
    const p = svg('path');
    p.setAttribute('d', path);
    const open = !custGeom;
    p.setAttribute('fill', open ? 'none' : (fill?.solid ?? (fill?.css === 'transparent' || !fill ? 'none' : fill.css)));
    if (line) strokeSvg(p, line, root);
    else p.setAttribute('stroke', 'none');
    root.append(p);
    box.append(root);
    return;
  }
  if (!fill && !line) return;

  const g = el('div', 'pptx-geom');
  if (fill?.image) {
    g.dataset.bg = fill.image;
    g.style.backgroundSize = fill.tile ? 'auto' : '100% 100%';
    if (fill.tile) g.style.backgroundRepeat = 'repeat';
  } else if (fill) {
    g.style.background = fill.css;
  }
  if (line) g.style.border = borderCss(line);
  if (prst === 'ellipse') g.style.borderRadius = '50%';
  else if (prst === 'roundRect') g.style.borderRadius = `${(adj(prstGeom, 'adj', 16667) / 100000) * Math.min(r.w, r.h)}px`;
  else {
    const pts = polygon(prst, r.w, r.h, prstGeom);
    if (pts) g.style.clipPath = `polygon(${pts.map(([x, y]) => `${x}px ${y}px`).join(', ')})`;
  }
  if (flip) g.style.transform = flip;
  box.append(g);
}

// ---------- 文字 ----------

interface TextStyle {
  /** bodyPr，优先级从高到低 */
  bodies: Element[];
  /** lstStyle / txStyles，优先级从高到低；每个里面按 lvlNpPr 取段落样式 */
  lists: Element[];
  /** lists 的前几项属于形状自身；之后才轮到形状样式（fontRef）给的颜色和字体 */
  own: number;
  refColor?: string;
  refFont?: string;
  refBold?: boolean;
}

function typeface(name: string | null | undefined, d: DrawCtx): string {
  if (!name) return '';
  const m = /^\+(mj|mn)-(lt|ea|cs)$/.exec(name);
  if (!m) return name;
  const f = m[1] === 'mj' ? d.theme.fonts.major : d.theme.fonts.minor;
  return m[2] === 'ea' ? f.ea : f.latin;
}

const truthy = (v: string | null | undefined) => v === '1' || v === 'true' || v === 'on';

const NUM_FORMATS: Record<string, (n: number) => string> = {
  arabic: String,
  alphaLc: (n) => alpha(n).toLowerCase(),
  alphaUc: alpha,
  romanLc: (n) => roman(n).toLowerCase(),
  romanUc: roman,
  circleNumDb: (n) => (n >= 1 && n <= 20 ? String.fromCharCode(0x2460 + n - 1) : String(n)),
  circleNumWd: (n) => (n >= 1 && n <= 20 ? String.fromCharCode(0x2460 + n - 1) : String(n)),
  ea1Chs: chinese,
  ea1Cht: chinese,
  ea1JpnChs: chinese,
};
function alpha(n: number): string {
  let s = '';
  for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}
function roman(n: number): string {
  const table: [number, string][] = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];
  let s = '';
  for (const [v, r] of table) for (; n >= v; n -= v) s += r;
  return s;
}
function chinese(n: number): string {
  const d = '〇一二三四五六七八九';
  if (n < 10) return d[n];
  if (n < 20) return `十${n % 10 ? d[n % 10] : ''}`;
  if (n < 100) return `${d[Math.floor(n / 10)]}十${n % 10 ? d[n % 10] : ''}`;
  return String(n);
}

function autoNumber(type: string, n: number): string {
  const m = /^(arabic|alphaLc|alphaUc|romanLc|romanUc|circleNumDb|circleNumWd|ea1Chs|ea1Cht|ea1JpnChs)(.*)$/.exec(type);
  const body = (NUM_FORMATS[m?.[1] ?? 'arabic'] ?? String)(n);
  const tail = m?.[2] ?? 'Period';
  if (tail.startsWith('ParenBoth')) return `(${body})`;
  if (tail.startsWith('ParenR')) return `${body})`;
  if (tail.startsWith('Plain') || tail === '') return body;
  return `${body}.`;
}

/** 一个文本框（或表格单元格）里的全部段落 */
function drawParagraphs(txBody: Element, style: TextStyle, d: DrawCtx, part: Part, out: HTMLElement) {
  const auto = kid(style.bodies[0], 'normAutofit');
  const scale = num(auto, 'fontScale', 100000) / 100000;
  const reduce = num(auto, 'lnSpcReduction', 0) / 100000;
  const counters: number[] = [];

  for (const p of kids(txBody, 'p')) {
    const pPr = kid(p, 'pPr');
    const lvl = Math.min(8, num(pPr, 'lvl', 0));
    const levels = style.lists.map((l) => kid(l, `lvl${lvl + 1}pPr`) ?? (l.localName === 'defaultTextStyle' ? kid(l, 'defPPr') : null));
    const ps = [pPr, ...levels].filter((n): n is Element => !!n);
    const pa = (name: string) => firstOf(ps, (n) => n.getAttribute(name));
    const pc = (name: string) => firstOf(ps, (n) => kid(n, name));
    // 段落、各级样式里的 defRPr；前面几项算形状自身
    const ownDefs = [pPr, ...levels.slice(0, style.own)].map((n) => kid(n, 'defRPr')).filter((n): n is Element => !!n);
    const restDefs = levels.slice(style.own).map((n) => kid(n, 'defRPr')).filter((n): n is Element => !!n);

    const runStyle = (rPr: Element | null) => {
      const own = [rPr, ...ownDefs].filter((n): n is Element => !!n);
      const all = [...own, ...restDefs];
      const attr = (name: string) => firstOf(all, (n) => n.getAttribute(name));
      const fillOf = (list: Element[]) => firstOf(list, (n) => paintOf(n, d));
      const fill = fillOf(own) ?? (style.refColor ? { css: style.refColor, solid: style.refColor } : undefined) ?? fillOf(restDefs);
      const font = (tag: string) => firstOf(own, (n) => kid(n, tag)?.getAttribute('typeface')) ?? (tag === 'latin' ? style.refFont : undefined) ?? firstOf(restDefs, (n) => kid(n, tag)?.getAttribute('typeface'));
      const latin = typeface(font('latin'), d) || d.theme.fonts.minor.latin;
      const ea = typeface(font('ea'), d) || d.theme.fonts.minor.ea;
      return {
        size: (Number(attr('sz') ?? 1800) / 100) * (4 / 3) * scale,
        bold: attr('b') != null ? truthy(attr('b')) : !!style.refBold,
        italic: truthy(attr('i')),
        underline: (attr('u') ?? 'none') !== 'none',
        strike: (attr('strike') ?? 'noStrike') !== 'noStrike',
        baseline: Number(attr('baseline') ?? 0),
        cap: attr('cap'),
        spacing: Number(attr('spc') ?? 0),
        color: fill?.solid ?? (fill?.css === 'transparent' ? 'transparent' : undefined) ?? 'rgb(0 0 0)',
        family: expand([latin, ea].filter(Boolean).join(', ') || 'sans-serif'),
        highlight: firstOf(own, (n) => colorOf(kid(n, 'highlight'), d)),
      };
    };
    type RunStyle = ReturnType<typeof runStyle>;
    const applyRun = (span: HTMLElement, s: RunStyle) => {
      const st = span.style;
      st.fontSize = `${s.size}px`;
      st.fontFamily = s.family;
      st.color = s.color;
      if (s.bold) st.fontWeight = '700';
      if (s.italic) st.fontStyle = 'italic';
      const deco = [s.underline && 'underline', s.strike && 'line-through'].filter(Boolean).join(' ');
      if (deco) st.textDecoration = deco;
      if (s.baseline) {
        st.verticalAlign = s.baseline > 0 ? 'super' : 'sub';
        st.fontSize = `${s.size * 0.7}px`;
      }
      if (s.cap === 'all') st.textTransform = 'uppercase';
      else if (s.cap === 'small') st.fontVariant = 'small-caps';
      if (s.spacing) st.letterSpacing = `${(s.spacing / 100) * (4 / 3)}px`;
      if (s.highlight) st.backgroundColor = s.highlight;
    };

    const para = el('p', 'pptx-p');
    const runs = kids(p).filter((k) => k.localName === 'r' || k.localName === 'br' || k.localName === 'fld');
    const hasText = runs.some((r) => r.localName !== 'br' && (kid(r, 't')?.textContent ?? '') !== '');
    const firstStyle = runStyle(kid(runs.find((r) => r.localName !== 'br') ?? null, 'rPr') ?? kid(p, 'endParaRPr'));

    const algn = pa('algn');
    if (algn) para.style.textAlign = { l: 'left', ctr: 'center', r: 'right', just: 'justify', dist: 'justify', thaiDist: 'justify', justLow: 'justify' }[algn] ?? 'left';
    const marL = Number(pa('marL') ?? 0) / EMU;
    const indent = Number(pa('indent') ?? 0) / EMU;
    if (marL) para.style.paddingLeft = `${marL}px`;
    if (indent) para.style.textIndent = `${indent}px`;
    const spacing = (name: string) => {
      const s = pc(name);
      const pts = num(kid(s, 'spcPts'), 'val');
      if (pts != null) return { px: (pts / 100) * (4 / 3) };
      const pct = num(kid(s, 'spcPct'), 'val');
      return pct != null ? { pct: pct / 100000 } : null;
    };
    const ln = spacing('lnSpc');
    if (ln?.px != null) para.style.lineHeight = `${ln.px * scale}px`;
    else para.style.lineHeight = String(1.2 * Math.max(0.1, (ln?.pct ?? 1) - reduce));
    for (const [name, prop] of [['spcBef', 'marginTop'], ['spcAft', 'marginBottom']] as const) {
      const s = spacing(name);
      if (s) para.style[prop] = `${s.px ?? (s.pct ?? 0) * 1.2 * firstStyle.size}px`;
    }

    // 项目符号和自动编号
    const bullet = hasText ? firstOf(ps, (n) => kids(n).find((k) => ['buNone', 'buChar', 'buAutoNum', 'buBlip'].includes(k.localName))) : undefined;
    let label = '';
    if (bullet?.localName === 'buAutoNum') {
      counters.length = lvl + 1;
      counters[lvl] = (counters[lvl] ?? num(bullet, 'startAt', 1) - 1) + 1;
      label = autoNumber(bullet.getAttribute('type') ?? 'arabicPeriod', counters[lvl]);
    } else {
      if (hasText) counters.length = lvl;
      if (bullet?.localName === 'buChar') label = bullet.getAttribute('char') ?? '•';
      else if (bullet?.localName === 'buBlip') label = '•';
    }
    if (label) {
      const b = el('span', 'pptx-bullet', label);
      const buFont = typeface(pc('buFont')?.getAttribute('typeface'), d);
      const sizePct = num(pc('buSzPct'), 'val');
      const sizePts = num(pc('buSzPts'), 'val');
      applyRun(b, {
        ...firstStyle,
        underline: false,
        strike: false,
        size: sizePts != null ? (sizePts / 100) * (4 / 3) * scale : firstStyle.size * (sizePct != null ? sizePct / 100000 : 1),
        color: colorOf(pc('buClr'), d) ?? firstStyle.color,
        // Wingdings 之类的符号字体在别的系统上没有，用正文字体显示常见的替代字符
        family: buFont && !/wingdings|symbol|webdings/i.test(buFont) ? expand(buFont) : firstStyle.family,
      });
      if (/wingdings/i.test(buFont)) b.textContent = { 'Ø': '➢', 'ü': '✓', 'l': '●', 'n': '■', 'q': '❑', 'v': '❖', '§': '■', 'Ü': '➢' }[label] ?? '•';
      if (indent < 0) b.style.minWidth = `${-indent}px`;
      else b.style.marginRight = '0.4em';
      para.append(b);
    }

    for (const r of runs) {
      if (r.localName === 'br') {
        para.append(el('br'));
        continue;
      }
      const rPr = kid(r, 'rPr');
      const text = r.localName === 'fld' && r.getAttribute('type') === 'slidenum' ? String(d.num) : (kid(r, 't')?.textContent ?? '');
      if (!text) continue;
      const rel = part.rels.get(nsAttr(kid(rPr, 'hlinkClick'), 'id') ?? '');
      const span = rel?.external && /^(https?|mailto):/i.test(rel.target) ? el('a') : el('span');
      if (span instanceof HTMLAnchorElement) {
        span.href = rel!.target;
        span.target = '_blank';
        span.rel = 'noopener noreferrer';
      }
      span.textContent = text;
      applyRun(span, runStyle(rPr));
      para.append(span);
    }
    if (!hasText) {
      // 空段落也占一行高度
      const blank = el('span', '', '​');
      applyRun(blank, runStyle(kid(p, 'endParaRPr')));
      para.append(blank);
    }
    out.append(para);
  }
}

/** 文本框：内边距、垂直对齐、竖排、分栏 */
function drawTextBox(txBody: Element, style: TextStyle, d: DrawCtx, part: Part, box: HTMLElement) {
  if (!kids(txBody, 'p').some((p) => kids(p).some((r) => (r.localName === 'r' || r.localName === 'fld') && kid(r, 't')?.textContent))) return;
  const attr = (name: string) => firstOf(style.bodies, (n) => n.getAttribute(name));
  const ins = (name: string, def: number) => Number(attr(name) ?? def) / EMU;
  const text = el('div', 'pptx-text');
  text.style.padding = `${ins('tIns', 45720)}px ${ins('rIns', 91440)}px ${ins('bIns', 45720)}px ${ins('lIns', 91440)}px`;
  text.style.justifyContent = { ctr: 'center', b: 'flex-end' }[attr('anchor') ?? 't'] ?? 'flex-start';
  if (attr('wrap') === 'none') text.style.whiteSpace = 'pre';
  const vert = attr('vert');
  if (vert && vert !== 'horz') text.style.writingMode = vert === 'vert270' ? 'sideways-lr' : 'vertical-rl';
  const cols = Number(attr('numCol') ?? 1);
  const inner = cols > 1 ? el('div') : text;
  if (cols > 1) {
    inner.style.columnCount = String(cols);
    inner.style.columnGap = `${Number(attr('spcCol') ?? 0) / EMU}px`;
    text.append(inner);
  }
  drawParagraphs(txBody, style, d, part, inner);
  box.append(text);
}

/** 占位符按类型套用母版的标题 / 正文 / 其他样式；普通文本框用演示文稿的默认样式 */
function baseLists(ph: Ph | null, d: DrawCtx): Element[] {
  const tx = kid(d.slide.master?.root, 'txStyles');
  if (!ph) return [d.deck.defaultText, kid(tx, 'otherStyle')].filter((n): n is Element => !!n);
  const t = masterType(ph.type);
  const name = t === 'title' ? 'titleStyle' : t === 'body' ? 'bodyStyle' : 'otherStyle';
  return [kid(tx, name)].filter((n): n is Element => !!n);
}

function plainText(txBody: Element | null): string[] {
  return kids(txBody, 'p').map((p) =>
    kids(p)
      .map((r) => (r.localName === 'br' ? ' ' : (kid(r, 't')?.textContent ?? '')))
      .join('')
      .trim(),
  );
}

// ---------- 各类元素 ----------

function shapeFillLine(chain: Element[], node: Element, d: DrawCtx, rels: Part['rels']) {
  const spPrs = chain.map((n) => kid(n, 'spPr')).filter((n): n is Element => !!n);
  const style = kid(node, 'style');
  const fillRef = refStyle(kid(style, 'fillRef'), 'fills', d);
  const fill = firstOf(spPrs, (n) => paintOf(n, d, rels)) ?? (fillRef ? paintFrom(fillRef.node, fillRef.ctx, rels) : undefined);
  const lnRef = refStyle(kid(style, 'lnRef'), 'lines', d);
  const refLine = lnRef ? lineOf(lnRef.node, lnRef.ctx) : undefined;
  const ownLn = firstOf(spPrs, (n) => kid(n, 'ln'));
  let line = lineOf(ownLn, d);
  // 只写了线宽、箭头没写颜色时，颜色取样式里的
  if (line === undefined && ownLn && refLine) {
    line = { ...refLine, width: ownLn.hasAttribute('w') ? num(ownLn, 'w', 12700) / EMU : refLine.width };
    const tail = kid(ownLn, 'tailEnd')?.getAttribute('type');
    const head = kid(ownLn, 'headEnd')?.getAttribute('type');
    if (tail) line.tail = tail === 'none' ? undefined : tail;
    if (head) line.head = head === 'none' ? undefined : head;
  }
  return { fill, line: line === undefined ? refLine : line, style };
}

function drawShape(node: Element, part: Part, d: DrawCtx, map: Mapper, out: HTMLElement) {
  const { chain, ph } = chainOf(node, part, d);
  const xf = firstOf(chain, (n) => xfrmOf(dig(n, 'spPr', 'xfrm')));
  if (!xf) return;
  const r = map(xf.x, xf.y, xf.w, xf.h);
  const box = el('div', 'pptx-shape');
  place(box, r, xf.rot);
  const { fill, line, style } = shapeFillLine(chain, node, d, part.rels);
  const geomSpPr = firstOf(chain, (n) => {
    const s = kid(n, 'spPr');
    return kid(s, 'prstGeom') || kid(s, 'custGeom') ? s : null;
  }) ?? null;
  drawGeom(box, geomSpPr, r, xf, fill, line);

  const txBody = kid(node, 'txBody');
  if (txBody) {
    const fontRef = kid(style, 'fontRef');
    const lists = chain.map((n) => dig(n, 'txBody', 'lstStyle')).filter((n): n is Element => !!n);
    drawTextBox(
      txBody,
      {
        bodies: chain.map((n) => dig(n, 'txBody', 'bodyPr')).filter((n): n is Element => !!n),
        lists: [...lists, ...baseLists(ph, d)],
        own: kid(txBody, 'lstStyle') ? 1 : 0,
        refColor: colorOf(fontRef, d) ?? undefined,
        refFont: fontRef ? (fontRef.getAttribute('idx') === 'major' ? '+mj-lt' : '+mn-lt') : undefined,
      },
      d,
      part,
      box,
    );
  }
  out.append(box);
}

function drawPicture(node: Element, part: Part, d: DrawCtx, map: Mapper, out: HTMLElement) {
  const { chain } = chainOf(node, part, d);
  const xf = firstOf(chain, (n) => xfrmOf(dig(n, 'spPr', 'xfrm')));
  if (!xf) return;
  const r = map(xf.x, xf.y, xf.w, xf.h);
  const box = el('div', 'pptx-pic');
  place(box, r, xf.rot);
  const blipFill = kid(node, 'blipFill');
  const rel = part.rels.get(nsAttr(kid(blipFill, 'blip'), 'embed') ?? '');
  if (rel && !rel.external) {
    const crop = kid(blipFill, 'srcRect');
    const [l, t, rr, b] = ['l', 't', 'r', 'b'].map((k) => num(crop, k, 0) / 100000);
    const w = r.w / Math.max(0.01, 1 - l - rr);
    const h = r.h / Math.max(0.01, 1 - t - b);
    const img = el('img');
    img.alt = '';
    img.draggable = false;
    img.dataset.src = rel.target;
    Object.assign(img.style, { left: `${-l * w}px`, top: `${-t * h}px`, width: `${w}px`, height: `${h}px` });
    const flip = `${xf.flipH ? 'scaleX(-1) ' : ''}${xf.flipV ? 'scaleY(-1)' : ''}`.trim();
    if (flip) img.style.transform = flip;
    box.append(img);
  }
  const spPr = kid(node, 'spPr');
  const prst = dig(spPr, 'prstGeom')?.getAttribute('prst');
  if (prst === 'ellipse') box.style.borderRadius = '50%';
  else if (prst === 'roundRect') box.style.borderRadius = `${(adj(kid(spPr, 'prstGeom'), 'adj', 16667) / 100000) * Math.min(r.w, r.h)}px`;
  const { line } = shapeFillLine(chain, node, d, part.rels);
  if (line) box.style.outline = borderCss(line);
  out.append(box);
}

async function drawGroup(node: Element, part: Part, d: DrawCtx, map: Mapper, out: HTMLElement) {
  const xf = xfrmOf(dig(node, 'grpSpPr', 'xfrm'));
  if (!xf) return drawTree(node, part, d, map, out, false);
  const r = map(xf.x, xf.y, xf.w, xf.h);
  const g = el('div', 'pptx-group');
  place(g, r, xf.rot);
  const ch = xf.ch ?? { x: xf.x, y: xf.y, w: xf.w, h: xf.h };
  const sx = ch.w ? r.w / ch.w : 1 / EMU;
  const sy = ch.h ? r.h / ch.h : 1 / EMU;
  await drawTree(node, part, d, (x, y, w, h) => ({ x: (x - ch.x) * sx, y: (y - ch.y) * sy, w: w * sx, h: h * sy }), g, false);
  out.append(g);
}

// PowerPoint 内置的表格样式不一定写进文件（很多生成工具只写一个 styleId）；
// 找不到定义时按最常用的“中度样式 2 - 强调 1”画，比完全没有边框和底色更接近原样
const BUILTIN_TABLE = (() => {
  const ln = (w: number) => `<a:ln w="${w}"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln>`;
  const fill = (tint?: number) => `<a:fill><a:solidFill><a:schemeClr val="accent1">${tint ? `<a:tint val="${tint}"/>` : ''}</a:schemeClr></a:solidFill></a:fill>`;
  const strong = '<a:tcTxStyle b="on"><a:schemeClr val="lt1"/></a:tcTxStyle>';
  const sides = ['left', 'right', 'top', 'bottom', 'insideH', 'insideV'].map((s) => `<a:${s}>${ln(12700)}</a:${s}>`).join('');
  let doc: Document | null = null;
  return () => {
    doc ??= new DOMParser().parseFromString(
      `<a:tblStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:wholeTbl><a:tcTxStyle><a:schemeClr val="dk1"/></a:tcTxStyle><a:tcStyle><a:tcBdr>${sides}</a:tcBdr>${fill(20000)}</a:tcStyle></a:wholeTbl>
        <a:band1H><a:tcStyle>${fill(40000)}</a:tcStyle></a:band1H>
        <a:band1V><a:tcStyle>${fill(40000)}</a:tcStyle></a:band1V>
        <a:firstCol>${strong}<a:tcStyle>${fill()}</a:tcStyle></a:firstCol>
        <a:lastCol>${strong}<a:tcStyle>${fill()}</a:tcStyle></a:lastCol>
        <a:lastRow>${strong}<a:tcStyle><a:tcBdr><a:top>${ln(38100)}</a:top></a:tcBdr>${fill()}</a:tcStyle></a:lastRow>
        <a:firstRow>${strong}<a:tcStyle><a:tcBdr><a:bottom>${ln(38100)}</a:bottom></a:tcBdr>${fill()}</a:tcStyle></a:firstRow>
      </a:tblStyle>`,
      'application/xml',
    );
    return doc.documentElement;
  };
})();

/** 表格样式（tableStyles.xml）里适用于某个单元格的各部分，后面的覆盖前面的 */
function tableParts(style: Element | undefined, tblPr: Element | null, row: number, col: number, rows: number, cols: number): Element[] {
  if (!style) return [];
  const on = (name: string) => truthy(tblPr?.getAttribute(name));
  const names = ['wholeTbl'];
  const bodyRow = row - (on('firstRow') ? 1 : 0);
  if (on('bandRow') && bodyRow >= 0) names.push(bodyRow % 2 ? 'band2H' : 'band1H');
  const bodyCol = col - (on('firstCol') ? 1 : 0);
  if (on('bandCol') && bodyCol >= 0) names.push(bodyCol % 2 ? 'band2V' : 'band1V');
  if (on('firstCol') && col === 0) names.push('firstCol');
  if (on('lastCol') && col === cols - 1) names.push('lastCol');
  if (on('lastRow') && row === rows - 1) names.push('lastRow');
  if (on('firstRow') && row === 0) names.push('firstRow');
  return names.map((n) => kid(style, n)).filter((n): n is Element => !!n);
}

function drawTable(tbl: Element, r: Rect, xf: Xfrm, part: Part, d: DrawCtx, box: HTMLElement) {
  const sx = xf.w ? r.w / xf.w : 1 / EMU;
  const sy = xf.h ? r.h / xf.h : 1 / EMU;
  const tblPr = kid(tbl, 'tblPr');
  const styleId = kid(tblPr, 'tableStyleId')?.textContent?.trim() ?? d.deck.defaultTableStyle;
  const style = d.deck.tableStyles.get(styleId) ?? (styleId ? BUILTIN_TABLE() : undefined);
  const widths = kids(kid(tbl, 'tblGrid'), 'gridCol').map((c) => num(c, 'w', 0) * sx);
  const rows = kids(tbl, 'tr');

  const table = el('table', 'pptx-table');
  table.style.width = `${widths.reduce((a, b) => a + b, 0)}px`;
  const colgroup = el('colgroup');
  for (const w of widths) {
    const c = el('col');
    c.style.width = `${w}px`;
    colgroup.append(c);
  }
  table.append(colgroup);

  const lineFrom = (holder: Element | null) => {
    const ref = kid(holder, 'lnRef');
    if (ref) {
      const s = refStyle(ref, 'lines', d);
      return s ? lineOf(s.node, s.ctx) : undefined;
    }
    return lineOf(kid(holder, 'ln'), d);
  };

  rows.forEach((tr, ri) => {
    const row = el('tr');
    row.style.height = `${num(tr, 'h', 0) * sy}px`;
    let ci = 0;
    for (const tc of kids(tr, 'tc')) {
      const col = ci;
      const span = num(tc, 'gridSpan', 1);
      ci += 1;
      if (truthy(tc.getAttribute('hMerge')) || truthy(tc.getAttribute('vMerge'))) continue;
      const rowSpan = num(tc, 'rowSpan', 1);
      const td = el('td');
      if (span > 1) td.colSpan = span;
      if (rowSpan > 1) td.rowSpan = rowSpan;
      const tcPr = kid(tc, 'tcPr');
      const styleParts = tableParts(style, tblPr, ri, col, rows.length, widths.length);

      const fill = paintOf(tcPr, d, part.rels) ?? firstOf([...styleParts].reverse(), (p) => {
        const f = dig(p, 'tcStyle', 'fill');
        const ref = refStyle(dig(p, 'tcStyle', 'fillRef'), 'fills', d);
        return f ? paintOf(f, d) : ref ? paintFrom(ref.node, ref.ctx) : undefined;
      });
      if (fill?.css) td.style.background = fill.css;

      const lastRow = ri + rowSpan >= rows.length;
      const lastCol = col + span >= widths.length;
      const sides = [
        ['lnT', ri === 0 ? 'top' : 'insideH', 'borderTop'],
        ['lnB', lastRow ? 'bottom' : 'insideH', 'borderBottom'],
        ['lnL', col === 0 ? 'left' : 'insideV', 'borderLeft'],
        ['lnR', lastCol ? 'right' : 'insideV', 'borderRight'],
      ] as const;
      for (const [own, name, prop] of sides) {
        const line = lineOf(kid(tcPr, own), d) ?? firstOf([...styleParts].reverse(), (p) => lineFrom(dig(p, 'tcStyle', 'tcBdr', name)));
        if (line) td.style[prop] = borderCss(line);
      }

      const ins = (name: string, def: number) => num(tcPr, name, def) / EMU;
      td.style.padding = `${ins('marT', 45720)}px ${ins('marR', 91440)}px ${ins('marB', 45720)}px ${ins('marL', 91440)}px`;
      td.style.verticalAlign = { ctr: 'middle', b: 'bottom' }[tcPr?.getAttribute('anchor') ?? 't'] ?? 'top';

      const txStyle = firstOf([...styleParts].reverse(), (p) => kid(p, 'tcTxStyle'));
      const txBody = kid(tc, 'txBody');
      if (txBody) {
        drawParagraphs(
          txBody,
          {
            bodies: [kid(txBody, 'bodyPr')].filter((n): n is Element => !!n),
            lists: [kid(txBody, 'lstStyle'), d.deck.defaultText].filter((n): n is Element => !!n),
            own: kid(txBody, 'lstStyle') ? 1 : 0,
            refColor: colorOf(txStyle, d) ?? colorOf(kid(txStyle, 'fontRef'), d) ?? undefined,
            refBold: truthy(txStyle?.getAttribute('b')),
          },
          d,
          part,
          td,
        );
      }
      row.append(td);
    }
    table.append(row);
  });
  box.append(table);
}

async function drawFrame(node: Element, part: Part, d: DrawCtx, map: Mapper, out: HTMLElement) {
  const xf = xfrmOf(kid(node, 'xfrm'));
  if (!xf) return;
  const r = map(xf.x, xf.y, xf.w, xf.h);
  const data = dig(node, 'graphic', 'graphicData');
  const uri = data?.getAttribute('uri') ?? '';
  const tbl = kid(data, 'tbl');
  if (tbl) {
    const box = el('div', 'pptx-shape pptx-table-box');
    place(box, r, xf.rot);
    drawTable(tbl, r, xf, part, d, box);
    out.append(box);
    return;
  }
  if (uri.endsWith('/diagram')) {
    // SmartArt：PowerPoint 会把排好版的形状存在 diagrams/drawingN.xml 里，直接画它
    const dm = part.rels.get(nsAttr(kid(data, 'relIds'), 'dm') ?? '');
    const dataDoc = dm && (await d.deck.pkg.xml(dm.target));
    const drawingRel = part.rels.get(find(dataDoc, 'dataModelExt')?.getAttribute('relId') ?? '');
    const drawing = drawingRel && (await loadPart(d.deck.pkg, drawingRel.target));
    const tree = drawing && find(drawing.doc, 'spTree');
    if (drawing && tree) {
      const g = el('div', 'pptx-group');
      place(g, r, xf.rot);
      const sx = xf.w ? r.w / xf.w : 1 / EMU;
      const sy = xf.h ? r.h / xf.h : 1 / EMU;
      await drawTree(tree, drawing, d, (x, y, w, h) => ({ x: x * sx, y: y * sy, w: w * sx, h: h * sy }), g, false);
      out.append(g);
      return;
    }
  }
  if (uri.endsWith('/chart') || uri.endsWith('/diagram') || uri.includes('chartex')) {
    const box = el('div', 'pptx-shape pptx-placeholder', uri.endsWith('/diagram') ? 'SmartArt' : '图表');
    place(box, r, xf.rot);
    out.append(box);
  }
}

async function drawTree(tree: Element | null, part: Part, d: DrawCtx, map: Mapper, out: HTMLElement, skipPh: boolean) {
  for (const node of kids(tree)) {
    if (node.localName === 'AlternateContent') {
      // 新格式放在 Choice 里，兼容的写法放在 Fallback 里
      const alt = kid(node, 'Fallback') ?? kid(node, 'Choice');
      if (alt) await drawTree(alt, part, d, map, out, skipPh);
      continue;
    }
    // 版式、母版上的占位符只是提示文字和位置，不画
    if (skipPh && phOf(node)) continue;
    switch (node.localName) {
      case 'sp':
      case 'cxnSp':
        drawShape(node, part, d, map, out);
        break;
      case 'pic':
        drawPicture(node, part, d, map, out);
        break;
      case 'grpSp':
        await drawGroup(node, part, d, map, out);
        break;
      case 'graphicFrame':
        await drawFrame(node, part, d, map, out);
        break;
    }
  }
}

function background(d: DrawCtx): Paint | undefined {
  for (const p of [d.slide.part, d.slide.layout, d.slide.master]) {
    const bg = dig(p?.root, 'cSld', 'bg');
    if (!p || !bg) continue;
    const pr = kid(bg, 'bgPr');
    const paint = pr ? paintOf(pr, d, p.rels) : undefined;
    if (paint) return paint;
    const ref = refStyle(kid(bg, 'bgRef'), 'fills', d);
    if (ref) return paintFrom(ref.node, ref.ctx, p.rels);
  }
  return undefined;
}

export interface DrawnSlide {
  el: HTMLElement;
  title: string;
  notes: string[];
}

export async function drawSlide(deck: Deck, slide: SlideRef, index: number): Promise<DrawnSlide> {
  const d: DrawCtx = { theme: slide.theme, map: slide.map, deck, slide, num: index + 1 };
  const root = el('div', 'pptx-slide');
  root.style.width = `${deck.width}px`;
  root.style.height = `${deck.height}px`;
  const bg = background(d);
  if (bg?.image) {
    root.dataset.bg = bg.image;
    root.style.backgroundSize = bg.tile ? 'auto' : 'cover';
    if (bg.tile) root.style.backgroundRepeat = 'repeat';
  } else if (bg?.css && bg.css !== 'transparent') {
    root.style.background = bg.css;
  }

  const showMaster = (p: Part | null) => p?.root.getAttribute('showMasterSp') !== '0';
  if (showMaster(slide.part)) {
    if (slide.master && showMaster(slide.layout)) await drawTree(spTree(slide.master), slide.master, d, rootMap, root, true);
    if (slide.layout) await drawTree(spTree(slide.layout), slide.layout, d, rootMap, root, true);
  }
  const tree = spTree(slide.part);
  await drawTree(tree, slide.part, d, rootMap, root, false);

  const titleShape = kids(tree).find((s) => isTitle(phOf(s)?.type ?? ''));
  const title = plainText(kid(titleShape, 'txBody')).filter(Boolean).join(' ');
  const notesBody = kids(spTree(slide.notes)).find((s) => phOf(s)?.type === 'body');
  return { el: root, title, notes: plainText(kid(notesBody, 'txBody')) };
}
