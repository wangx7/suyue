// 智能深色的核心配色：亮度反转、保留色相与饱和度（白纸→深灰，黑字→浅灰，红字仍是红色）

const PAPER_L = 0.078; // 白色映射到的亮度，对应背景 #131415
const INK_L = 0.716; // 黑色映射到的亮度，对应正文 #b5b5b8
const MIN_COLORED_TEXT_L = 0.6; // 彩色文字的最低亮度，保证在深色背景上看得清

export type Rgba = [number, number, number, number];

/** 解析 #rrggbb / rgb() / rgba()（浏览器规范化后的写法） */
export function parseColor(s: string): Rgba | null {
  const hex = /^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(s);
  if (hex) return [parseInt(hex[1], 16), parseInt(hex[2], 16), parseInt(hex[3], 16), 1];
  const m = /^rgba?\(([\d.]+),\s*([\d.]+),\s*([\d.]+)(?:,\s*([\d.]+))?\)$/.exec(s);
  return m ? [Number(m[1]), Number(m[2]), Number(m[3]), m[4] === undefined ? 1 : Number(m[4])] : null;
}

function rgbToHsl(r: number, g: number, b: number): [number, number, number] {
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  const h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [h / 6, s, l];
}

function hslToRgb(h: number, s: number, l: number): [number, number, number] {
  if (!s) return [l, l, l];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const f = (t: number) => {
    t = (t + 1) % 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  return [f(h + 1 / 3), f(h), f(h - 1 / 3)];
}

/**
 * 返回深色版本，写法为 rgb(r,g,b) 或 rgba(r,g,b,a)。
 * role 为 'text' 时额外提亮彩色文字（深红、深蓝直接反转后在深色背景上偏暗）。
 */
export function darkColor([r0, g0, b0, a]: Rgba, role: 'text' | 'fill' = 'fill'): string {
  const [h, s, l] = rgbToHsl(r0 / 255, g0 / 255, b0 / 255);
  let l2 = PAPER_L + (1 - l) * (INK_L - PAPER_L);
  if (role === 'text' && s > 0.15) l2 = Math.max(l2, MIN_COLORED_TEXT_L);
  const [r, g, b] = hslToRgb(h, s, l2).map((v) => Math.round(v * 255));
  return a === 1 ? `rgb(${r},${g},${b})` : `rgba(${r},${g},${b},${a})`;
}
