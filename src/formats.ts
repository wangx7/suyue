// 格式识别和渲染器加载，网页版和 VS Code 插件共用
import type { RecentItem } from './store';
import type { Renderer } from './renderers/types';

export type Kind = RecentItem['kind'];

// 每种格式的渲染器按需加载：打开 Word 时不会下载表格解析库
export const loaders: Record<Kind, () => Promise<Renderer>> = {
  docx: () => import('./renderers/docx').then((m) => m.renderDocx),
  doc: () => import('./renderers/doc').then((m) => m.renderDoc),
  sheet: () => import('./renderers/sheet').then((m) => m.renderSheet),
  md: () => import('./renderers/md').then((m) => m.renderMarkdown),
};
export const KIND_LABEL: Record<Kind, string> = { docx: 'Word 文档', doc: 'Word 97-2003 文档', sheet: '表格', md: 'Markdown' };

export const SHEET_EXTS = ['xlsx', 'xlsm', 'xlsb', 'xls', 'ods', 'csv', 'tsv'];
export const MD_EXTS = ['md', 'markdown', 'mdown', 'mkd', 'mdx'];

export function extOf(name: string): string {
  const m = /\.([a-z0-9]+)$/i.exec(name);
  return m ? m[1].toLowerCase() : '';
}

function bytesInclude(buf: ArrayBuffer, needle: string, fromEnd: boolean): boolean {
  const len = Math.min(buf.byteLength, 128 * 1024);
  const view = new Uint8Array(buf, fromEnd ? buf.byteLength - len : 0, len);
  return new TextDecoder('latin1').decode(view).includes(needle);
}

/** 先看文件头魔数，再看扩展名，扩展名缺失时再从 zip 目录 / OLE 目录里找特征 */
export function detect(name: string, buf: ArrayBuffer): Kind | null {
  const ext = extOf(name);
  const b = new Uint8Array(buf, 0, Math.min(8, buf.byteLength));
  if (b[0] === 0x50 && b[1] === 0x4b) {
    if (['docx', 'docm', 'dotx'].includes(ext)) return 'docx';
    if (SHEET_EXTS.includes(ext)) return 'sheet';
    if (bytesInclude(buf, 'word/', true)) return 'docx';
    if (bytesInclude(buf, 'xl/', true) || bytesInclude(buf, 'content.xml', true)) return 'sheet';
    return null;
  }
  if (b[0] === 0xd0 && b[1] === 0xcf && b[2] === 0x11 && b[3] === 0xe0) {
    if (ext === 'doc' || ext === 'dot') return 'doc';
    if (ext === 'xls') return 'sheet';
    return bytesInclude(buf, 'W\0o\0r\0d\0D\0o\0c', false) ? 'doc' : 'sheet';
  }
  if (MD_EXTS.includes(ext)) return 'md';
  if (SHEET_EXTS.includes(ext) || ext === 'txt') return 'sheet';
  return null;
}

export function sizeText(n: number): string {
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(0)} KB`;
  return `${(n / 1024 / 1024).toFixed(1)} MB`;
}
