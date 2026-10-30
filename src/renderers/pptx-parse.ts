import { dig, kid, kids, num, nsAttr, Pkg, resolvePath } from './zip';

// PPTX 的包结构、主题和颜色。幻灯片 → 版式 → 母版 → 主题沿关系文件串起来；
// 颜色可能是主题色加亮度/透明度修饰，填充可能引用主题的样式列表，这里都解析成 CSS 能用的值。

export const EMU = 9525; // 1px（96dpi）对应的 EMU

export interface Rel {
  type: string;
  target: string;
  external: boolean;
}
export type Rels = Map<string, Rel>;

export interface Part {
  path: string;
  doc: Document;
  root: Element;
  rels: Rels;
}

export interface Theme {
  colors: Record<string, string>;
  fonts: { major: { latin: string; ea: string }; minor: { latin: string; ea: string } };
  fills: Element[];
  bgFills: Element[];
  lines: Element[];
}

export interface SlideRef {
  part: Part;
  layout: Part | null;
  master: Part | null;
  theme: Theme;
  /** 主题色名的映射（bg1 → lt1 等），幻灯片可以覆盖母版的映射 */
  map: Record<string, string>;
  notes: Part | null;
  hidden: boolean;
}

export interface Deck {
  pkg: Pkg;
  width: number;
  height: number;
  slides: SlideRef[];
  defaultText: Element | null;
  tableStyles: Map<string, Element>;
  /** 没写 tableStyleId 的表格用的样式 */
  defaultTableStyle: string;
}

export const relType = (rel: Rel, suffix: string) => rel.type.endsWith(`/${suffix}`);

async function rels(pkg: Pkg, path: string): Promise<Rels> {
  const i = path.lastIndexOf('/');
  const doc = await pkg.xml(`${path.slice(0, i + 1)}_rels/${path.slice(i + 1)}.rels`);
  const map: Rels = new Map();
  for (const r of kids(doc?.documentElement)) {
    const id = r.getAttribute('Id');
    const target = r.getAttribute('Target') ?? '';
    const external = r.getAttribute('TargetMode') === 'External';
    if (id) map.set(id, { type: r.getAttribute('Type') ?? '', target: external ? target : resolvePath(path, target), external });
  }
  return map;
}

const parts = new WeakMap<Pkg, Map<string, Promise<Part | null>>>();

export function loadPart(pkg: Pkg, path: string): Promise<Part | null> {
  let cache = parts.get(pkg);
  if (!cache) parts.set(pkg, (cache = new Map()));
  let p = cache.get(path);
  if (!p) {
    p = Promise.all([pkg.xml(path), rels(pkg, path)]).then(([doc, r]) => (doc ? { path, doc, root: doc.documentElement, rels: r } : null));
    cache.set(path, p);
  }
  return p;
}

export function relTarget(part: Part, suffix: string): string | null {
  for (const r of part.rels.values()) if (!r.external && relType(r, suffix)) return r.target;
  return null;
}

function parseTheme(doc: Document | null): Theme {
  const root = doc?.documentElement;
  const el = dig(root, 'themeElements');
  const colors: Record<string, string> = {};
  for (const c of kids(dig(el, 'clrScheme'))) {
    const v = kid(c, 'srgbClr')?.getAttribute('val') ?? kid(c, 'sysClr')?.getAttribute('lastClr');
    if (v) colors[c.localName] = v;
  }
  const font = (name: string) => {
    const f = dig(el, 'fontScheme', name);
    return { latin: dig(f, 'latin')?.getAttribute('typeface') ?? '', ea: dig(f, 'ea')?.getAttribute('typeface') ?? '' };
  };
  const fmt = dig(el, 'fmtScheme');
  return {
    colors,
    fonts: { major: font('majorFont'), minor: font('minorFont') },
    fills: kids(kid(fmt, 'fillStyleLst')),
    bgFills: kids(kid(fmt, 'bgFillStyleLst')),
    lines: kids(kid(fmt, 'lnStyleLst')),
  };
}

const DEFAULT_MAP: Record<string, string> = { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' };

function colorMap(node: Element | null): Record<string, string> | null {
  if (!node) return null;
  const map: Record<string, string> = {};
  for (const a of node.attributes) map[a.name] = a.value;
  return map;
}

export async function loadDeck(data: ArrayBuffer): Promise<Deck> {
  const pkg = await Pkg.open(data);
  const pres = await loadPart(pkg, 'ppt/presentation.xml');
  if (!pres) throw new Error(pkg.has('EncryptedPackage') ? '演示文稿已加密，暂不支持' : '不是有效的 PPTX 文件');
  const size = kid(pres.root, 'sldSz');
  const themes = new Map<string, Promise<Theme>>();
  const theme = (path: string | null) => {
    const key = path ?? '';
    if (!themes.has(key)) themes.set(key, (path ? pkg.xml(path) : Promise.resolve(null)).then(parseTheme));
    return themes.get(key)!;
  };

  const slides = await Promise.all(
    kids(kid(pres.root, 'sldIdLst'), 'sldId').map(async (s): Promise<SlideRef | null> => {
      const rel = pres.rels.get(nsAttr(s, 'id') ?? '');
      const part = rel && (await loadPart(pkg, rel.target));
      if (!part) return null;
      const layoutPath = relTarget(part, 'slideLayout');
      const layout = layoutPath ? await loadPart(pkg, layoutPath) : null;
      const masterPath = layout && relTarget(layout, 'slideMaster');
      const master = masterPath ? await loadPart(pkg, masterPath) : null;
      const notesPath = relTarget(part, 'notesSlide');
      const override = (p: Part | null) => colorMap(dig(p?.root, 'clrMapOvr', 'overrideClrMapping'));
      return {
        part,
        layout,
        master,
        theme: await theme(master && relTarget(master, 'theme')),
        map: override(part) ?? override(layout) ?? colorMap(kid(master?.root, 'clrMap')) ?? DEFAULT_MAP,
        notes: notesPath ? await loadPart(pkg, notesPath) : null,
        hidden: part.root.getAttribute('show') === '0',
      };
    }),
  );

  const tableStyles = new Map<string, Element>();
  const ts = relTarget(pres, 'tableStyles');
  const tsDoc = ts ? await pkg.xml(ts) : null;
  const defaultTableStyle = tsDoc?.documentElement.getAttribute('def') ?? '';
  for (const s of kids(tsDoc?.documentElement, 'tblStyle')) {
    const id = s.getAttribute('styleId');
    if (id) tableStyles.set(id, s);
  }

  return {
    pkg,
    width: num(size, 'cx', 9144000) / EMU,
    height: num(size, 'cy', 6858000) / EMU,
    slides: slides.filter((s): s is SlideRef => !!s),
    defaultText: kid(pres.root, 'defaultTextStyle'),
    tableStyles,
    defaultTableStyle,
  };
}

// ---------- 颜色 ----------

type Rgba = [number, number, number, number];

export interface ColorCtx {
  theme: Theme;
  map: Record<string, string>;
  /** 样式引用里给出的占位色（phClr） */
  ph?: Rgba;
}

const COLOR_TAGS = new Set(['srgbClr', 'schemeClr', 'sysClr', 'prstClr', 'scrgbClr', 'hslClr']);
const PRESET: Record<string, string> = {
  black: '000000', white: 'ffffff', red: 'ff0000', green: '008000', blue: '0000ff', yellow: 'ffff00', gray: '808080',
  dkGray: 'a9a9a9', ltGray: 'd3d3d3', orange: 'ffa500', purple: '800080', navy: '000080', cyan: '00ffff', magenta: 'ff00ff',
};

const hex = (v: string): Rgba => [parseInt(v.slice(0, 2), 16), parseInt(v.slice(2, 4), 16), parseInt(v.slice(4, 6), 16), 1];

function toHsl([r, g, b]: Rgba): [number, number, number] {
  r /= 255;
  g /= 255;
  b /= 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  const h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [h / 6, s, l];
}

function fromHsl(h: number, s: number, l: number, a: number): Rgba {
  if (!s) return [l * 255, l * 255, l * 255, a];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const f = (t: number) => {
    t = (t + 1) % 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  return [f(h + 1 / 3) * 255, f(h) * 255, f(h - 1 / 3) * 255, a];
}

const clamp01 = (v: number) => Math.min(1, Math.max(0, v));

function baseColor(node: Element, c: ColorCtx): Rgba | null {
  const val = node.getAttribute('val') ?? '';
  switch (node.localName) {
    case 'srgbClr':
      return /^[0-9a-f]{6}$/i.test(val) ? hex(val) : null;
    case 'sysClr': {
      const v = node.getAttribute('lastClr') ?? (val === 'window' ? 'ffffff' : '000000');
      return hex(v);
    }
    case 'prstClr':
      return PRESET[val] ? hex(PRESET[val]) : null;
    case 'scrgbClr': {
      const ch = (n: string) => clamp01(num(node, n, 0) / 100000) ** (1 / 2.2) * 255;
      return [ch('r'), ch('g'), ch('b'), 1];
    }
    case 'hslClr':
      return fromHsl(num(node, 'hue', 0) / 21600000, num(node, 'sat', 0) / 100000, num(node, 'lum', 0) / 100000, 1);
    case 'schemeClr': {
      if (val === 'phClr') return c.ph ? [...c.ph] : null;
      const v = c.theme.colors[c.map[val] ?? val];
      return v ? hex(v) : null;
    }
  }
  return null;
}

/** 颜色元素（srgbClr 等）连同亮度、透明度修饰 */
function colorFrom(node: Element, c: ColorCtx): Rgba | null {
  let rgba = baseColor(node, c);
  if (!rgba) return null;
  for (const m of kids(node)) {
    const v = num(m, 'val', 0) / 100000;
    switch (m.localName) {
      case 'alpha':
        rgba[3] = v;
        break;
      case 'tint':
        rgba = [...rgba.slice(0, 3).map((x) => x + (255 - x) * (1 - v)), rgba[3]] as Rgba;
        break;
      case 'shade':
        rgba = [...rgba.slice(0, 3).map((x) => x * v), rgba[3]] as Rgba;
        break;
      case 'lumMod':
      case 'lumOff':
      case 'satMod': {
        const [h, s, l] = toHsl(rgba);
        if (m.localName === 'satMod') rgba = fromHsl(h, clamp01(s * v), l, rgba[3]);
        else rgba = fromHsl(h, s, clamp01(m.localName === 'lumMod' ? l * v : l + v), rgba[3]);
        break;
      }
    }
  }
  return rgba;
}

const css = ([r, g, b, a]: Rgba) => (a >= 1 ? `rgb(${Math.round(r)} ${Math.round(g)} ${Math.round(b)})` : `rgb(${Math.round(r)} ${Math.round(g)} ${Math.round(b)} / ${+a.toFixed(3)})`);

function rgbaOf(parent: Element | null | undefined, c: ColorCtx): Rgba | null {
  for (const k of kids(parent)) if (COLOR_TAGS.has(k.localName)) return colorFrom(k, c);
  return null;
}

/** parent 下的颜色元素转成 CSS 颜色 */
export function colorOf(parent: Element | null | undefined, c: ColorCtx): string | null {
  const rgba = rgbaOf(parent, c);
  return rgba && css(rgba);
}

// ---------- 填充和线条 ----------

export interface Paint {
  /** CSS background；图片填充时为空，见 image */
  css: string;
  /** 渐变或图片时的代表色，给文字、线条这类只能用纯色的地方 */
  solid?: string;
  /** 图片填充：包内路径 */
  image?: string;
  /** 图片填充是平铺 */
  tile?: boolean;
}

const FILL_TAGS = new Set(['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill']);

/** 单个填充元素（solidFill 等）；rels 用于解析图片填充 */
export function paintFrom(node: Element, c: ColorCtx, rels?: Rels): Paint | undefined {
  switch (node.localName) {
    case 'noFill':
      return { css: 'transparent' };
    case 'solidFill': {
      const v = colorOf(node, c);
      return v ? { css: v, solid: v } : undefined;
    }
    case 'pattFill': {
      const v = colorOf(kid(node, 'fgClr'), c);
      return v ? { css: v, solid: v } : undefined;
    }
    case 'gradFill': {
      const stops = kids(kid(node, 'gsLst'), 'gs')
        .map((g) => ({ pos: num(g, 'pos', 0) / 1000, color: colorOf(g, c) }))
        .filter((s): s is { pos: number; color: string } => !!s.color)
        .sort((a, b) => a.pos - b.pos);
      if (!stops.length) return undefined;
      const list = stops.map((s) => `${s.color} ${s.pos}%`).join(', ');
      const path = kid(node, 'path')?.getAttribute('path');
      const value = path && path !== 'linear' ? `radial-gradient(circle, ${list})` : `linear-gradient(${num(kid(node, 'lin'), 'ang', 0) / 60000 + 90}deg, ${list})`;
      return { css: stops.length > 1 ? value : stops[0].color, solid: stops[0].color };
    }
    case 'blipFill': {
      const rel = rels?.get(nsAttr(kid(node, 'blip'), 'embed') ?? '');
      return rel && !rel.external ? { css: '', image: rel.target, tile: !!kid(node, 'tile') } : undefined;
    }
  }
  return undefined;
}

/** 容器（spPr、bgPr、tcPr 等）里的填充；没写填充返回 undefined */
export function paintOf(parent: Element | null | undefined, c: ColorCtx, rels?: Rels): Paint | undefined {
  for (const k of kids(parent)) if (FILL_TAGS.has(k.localName)) return paintFrom(k, c, rels);
  return undefined;
}

/** 形状样式里的 fillRef / lnRef / bgRef：idx 指向主题的样式列表，子元素颜色作为 phClr */
export function refStyle(ref: Element | null, list: 'fills' | 'lines', c: ColorCtx): { node: Element; ctx: ColorCtx } | null {
  const idx = num(ref, 'idx', 0);
  if (!ref || !idx) return null;
  const arr = list === 'lines' ? c.theme.lines : idx >= 1001 ? c.theme.bgFills : c.theme.fills;
  const node = arr[(idx >= 1001 ? idx - 1001 : idx) - 1];
  return node ? { node, ctx: { ...c, ph: rgbaOf(ref, c) ?? undefined } } : null;
}

export interface Line {
  width: number;
  color: string;
  dash: string;
  head?: string;
  tail?: string;
}

/** a:ln；noFill 返回 null，没写返回 undefined */
export function lineOf(ln: Element | null | undefined, c: ColorCtx): Line | null | undefined {
  if (!ln) return undefined;
  const paint = paintOf(ln, c);
  if (paint?.css === 'transparent') return null;
  if (!paint) return undefined;
  const dash = kid(ln, 'prstDash')?.getAttribute('val') ?? 'solid';
  const end = (name: string) => {
    const t = kid(ln, name)?.getAttribute('type');
    return t && t !== 'none' ? t : undefined;
  };
  return {
    width: Math.max(0.75, num(ln, 'w', 12700) / EMU),
    color: paint.solid ?? paint.css,
    dash,
    head: end('headEnd'),
    tail: end('tailEnd'),
  };
}
