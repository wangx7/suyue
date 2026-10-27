// 字体回退：Word 文档里写的多是 Windows 字体（宋体、微软雅黑、Calibri…），在 macOS 等系统上没有时
// 浏览器会退回默认衬线体。给每个字体补上各平台的同类字体和通用族名，尽量保持字形风格。

type Generic = 'serif' | 'sans-serif' | 'monospace';

const FAMILIES: [names: string[], fallbacks: string[], generic: Generic][] = [
  [['宋体', 'SimSun', '新宋体', 'NSimSun', '宋体-简', '华文宋体', 'STSong', '方正书宋_GBK', '方正书宋简体'], ['SimSun', 'Songti SC', 'STSong'], 'serif'],
  [['方正小标宋简体', '方正小标宋_GBK', '小标宋', '华文中宋', 'STZhongsong'], ['STZhongsong', 'Songti SC'], 'serif'],
  [['思源宋体', 'Source Han Serif SC', 'Noto Serif SC', 'Noto Serif CJK SC'], ['Source Han Serif SC', 'Noto Serif SC', 'Songti SC'], 'serif'],
  [['仿宋', 'FangSong', '仿宋_GB2312', '华文仿宋', 'STFangsong', '方正仿宋_GBK', '方正仿宋简体'], ['FangSong', 'STFangsong'], 'serif'],
  [['楷体', 'KaiTi', '楷体_GB2312', '华文楷体', 'STKaiti', '方正楷体_GBK', '方正楷体简体'], ['KaiTi', 'Kaiti SC', 'STKaiti'], 'serif'],
  [['黑体', 'SimHei', '华文黑体', 'STHeiti', '方正黑体_GBK', '方正黑体简体'], ['SimHei', 'Heiti SC', 'STHeiti', 'PingFang SC'], 'sans-serif'],
  [['微软雅黑', 'Microsoft YaHei', 'Microsoft YaHei UI', '等线', 'DengXian', '等线 Light', 'DengXian Light', '思源黑体', 'Source Han Sans SC', 'Noto Sans SC', 'Noto Sans CJK SC'], ['Microsoft YaHei', 'PingFang SC', 'Hiragino Sans GB', 'Noto Sans CJK SC'], 'sans-serif'],
  [['Calibri', 'Calibri Light', 'Carlito', 'Aptos', 'Segoe UI', 'Tahoma', 'Verdana'], ['Carlito', 'Helvetica Neue', 'Arial'], 'sans-serif'],
  [['Arial', 'Helvetica', 'Helvetica Neue', 'Arial Unicode MS'], ['Helvetica Neue', 'Arial'], 'sans-serif'],
  [['Cambria', 'Caladea', 'Georgia', 'Book Antiqua', 'Palatino Linotype'], ['Caladea', 'Georgia', 'Times New Roman'], 'serif'],
  [['Times New Roman', 'Times', 'Liberation Serif'], ['Times New Roman', 'Times'], 'serif'],
  [['Consolas', 'Courier New', 'Courier', 'Menlo', 'Monaco', 'Lucida Console', 'Source Code Pro'], ['Consolas', 'SF Mono', 'Menlo', 'Monaco', 'Courier New'], 'monospace'],
];
const LOOKUP = new Map(FAMILIES.flatMap(([names, fallbacks, generic]) => names.map((n) => [n.toLowerCase(), { fallbacks, generic }])));
const GENERICS = new Set(['serif', 'sans-serif', 'monospace', 'cursive', 'fantasy', 'system-ui']);

const quote = (n: string) => (GENERICS.has(n) ? n : `"${n.replace(/"/g, '')}"`);

/** 展开 font-family 列表：每个字体后面紧跟它的替代字体，通用族名放最后 */
function expand(list: string): string {
  if (list.includes('var(')) return list;
  const out: string[] = [];
  let generic = '';
  for (const raw of list.split(',')) {
    const name = raw.trim().replace(/^(['"])(.*)\1$/, '$2');
    if (!name) continue;
    if (GENERICS.has(name)) {
      generic ||= name;
      continue;
    }
    const known = LOOKUP.get(name.toLowerCase());
    out.push(name, ...(known?.fallbacks ?? []));
    generic ||= known?.generic ?? '';
  }
  if (generic) out.push(generic);
  return [...new Set(out)].map(quote).join(', ');
}

const DECL = /(font-family|--docx-[\w-]+-font)(\s*:\s*)([^;}\r\n]+)/g;

/** 改写 docx-preview 生成的样式表（含主题字体变量）和行内 font-family，需在测量排版之前调用 */
export function fixFonts(root: HTMLElement) {
  for (const style of root.querySelectorAll('style')) {
    const text = style.textContent ?? '';
    if (text.includes('@font-face')) continue; // 内嵌字体的声明，族名必须是单个名字
    const next = text.replace(DECL, (_, prop: string, sep: string, value: string) => prop + sep + expand(value));
    if (next !== text) style.textContent = next;
  }
  for (const n of root.querySelectorAll<HTMLElement>('[style*="font-family"]')) {
    if (n.style.fontFamily) n.style.fontFamily = expand(n.style.fontFamily);
  }
}
