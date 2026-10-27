import * as CFB from 'cfb';
import { el } from '../ui';
import { cssZoom } from '../zoom';
import { trackHeadings, type Heading } from './outline';
import type { Renderer } from './types';

// 旧版 Word 97-2003 (.doc) 二进制格式：从 piece table 中提取正文纯文本。
// 参考 [MS-DOC] 2.5.1 Fib 与 2.9.38 Clx。

function streamBytes(cfb: CFB.CFB$Container, name: string): Uint8Array | null {
  const entry = CFB.find(cfb, name);
  return entry?.content ? new Uint8Array(entry.content as ArrayLike<number>) : null;
}

function extractText(buf: ArrayBuffer): string {
  const cfb = CFB.read(new Uint8Array(buf), { type: 'array' });
  const word = streamBytes(cfb, 'WordDocument');
  if (!word) throw new Error('不是有效的 Word 文档');
  const wv = new DataView(word.buffer, word.byteOffset, word.byteLength);

  if (wv.getUint16(0, true) !== 0xa5ec) throw new Error('不是有效的 Word 文档');
  if (wv.getUint16(2, true) < 0xc1) throw new Error('Word 95 及更早版本的 .doc 暂不支持');
  const flags = wv.getUint16(0x0a, true);
  if (flags & 0x0100) throw new Error('该 .doc 文件已加密，暂不支持');

  // Fib 各段长度可变，逐段定位
  const csw = wv.getUint16(32, true);
  const lwStart = 34 + csw * 2 + 2;
  const cslw = wv.getUint16(34 + csw * 2, true);
  const ccpText = wv.getUint32(lwStart + 12, true);
  const fcLcbStart = lwStart + cslw * 4 + 2;
  const fcClx = wv.getUint32(fcLcbStart + 33 * 8, true);
  const lcbClx = wv.getUint32(fcLcbStart + 33 * 8 + 4, true);

  const table = streamBytes(cfb, flags & 0x0200 ? '1Table' : '0Table');
  if (!table || !lcbClx) throw new Error('文档结构不完整');
  const tv = new DataView(table.buffer, table.byteOffset, table.byteLength);

  // Clx = 若干 Prc（跳过） + 一个 Pcdt（piece table）
  let pos = fcClx;
  while (table[pos] === 0x01) pos += 3 + tv.getInt16(pos + 1, true);
  if (table[pos] !== 0x02) throw new Error('无法解析文档文本');
  const lcb = tv.getUint32(pos + 1, true);
  const plc = pos + 5;
  const n = (lcb - 4) / 12;

  const utf16 = new TextDecoder('utf-16le');
  const cp1252 = new TextDecoder('windows-1252');
  let text = '';
  for (let i = 0; i < n && text.length < ccpText; i++) {
    const cpStart = tv.getUint32(plc + i * 4, true);
    const cpEnd = tv.getUint32(plc + (i + 1) * 4, true);
    const fc = tv.getUint32(plc + (n + 1) * 4 + i * 8 + 2, true);
    const len = cpEnd - cpStart;
    if (fc & 0x40000000) {
      const off = (fc & 0x3fffffff) / 2;
      text += cp1252.decode(word.subarray(off, off + len));
    } else {
      text += utf16.decode(word.subarray(fc, fc + len * 2));
    }
  }
  return cleanText(text.slice(0, ccpText));
}

/** 去掉域代码（保留域结果）和各种控制字符 */
function cleanText(raw: string): string {
  let out = '';
  const fields: boolean[] = []; // true = 处于域代码部分
  let inCode = 0;
  for (const ch of raw) {
    const code = ch.charCodeAt(0);
    if (code === 0x13) {
      fields.push(true);
      inCode++;
    } else if (code === 0x14) {
      if (fields[fields.length - 1]) {
        fields[fields.length - 1] = false;
        inCode--;
      }
    } else if (code === 0x15) {
      if (fields.pop()) inCode--;
    } else if (inCode) {
      continue;
    } else if (code === 0x07) {
      // 单元格结束标记；紧跟在单元格结束后的那个是行结束标记
      out += out.endsWith('\t') ? '\r' : '\t';
    } else if (code === 0x0b || code === 0x0c) out += '\n'; // 手动换行 / 分页
    else if (code === 0x1e) out += '-';
    else if (code < 0x20 && code !== 0x0d && code !== 0x09) continue; // 嵌入对象、脚注引用等
    else out += ch;
  }
  return out;
}

/** 编号形式的标题：第一章 / 第二节 / 一、/（一）/ 1. / 1.1 */
const HEADING_PATTERNS: [RegExp, number][] = [
  [/^第[一二三四五六七八九十百零〇\d]+[章篇部]/, 1],
  [/^第[一二三四五六七八九十百零〇\d]+[节条]/, 2],
  [/^[一二三四五六七八九十]+(?:、|[.．](?!\d))/, 1],
  [/^[（(][一二三四五六七八九十]+[)）]/, 2],
  [/^\d+\.\d+\.\d+\s*\S/, 3],
  [/^\d+\.\d+\s*\S/, 2],
  [/^\d+[.．、]\s*\S/, 2],
];

/** 旧格式没有样式信息，只能按编号推测标题：短、无句末标点、符合编号形式 */
function headingLevel(text: string): number {
  if (text.length > 40 || /[。．!?！？；;，,：:]$/.test(text)) return 0;
  for (const [re, level] of HEADING_PATTERNS) if (re.test(text)) return level;
  return 0;
}

export const renderDoc: Renderer = async (data, { host, shell, saved }) => {
  const text = extractText(data);
  const scroller = el('div', 'doc-scroll');
  const body = el('div', 'text-read');
  body.append(el('div', 'notice', '旧版 .doc 格式仅显示正文文字，排版和图片不保留。另存为 .docx 可获得完整效果。'));
  scroller.append(body);
  host.append(scroller);

  const headings: Heading[] = [];
  const frag = document.createDocumentFragment();
  text.split('\r').forEach((para, i) => {
    const t = para.trim();
    // 第一段若较短，通常是文档标题
    const level = i === 0 && t && t.length <= 40 ? 1 : headingLevel(t);
    const node = el(level ? (`h${Math.min(level + 1, 4)}` as 'h2') : 'p', '', para);
    if (level) headings.push({ el: node, item: { id: `h${headings.length}`, title: t, level } });
    frag.append(node);
  });
  body.append(frag);

  const tracker = trackHeadings(scroller, shell);
  // 只有一个“标题”时多半是误判，不显示目录
  tracker.set(headings.length >= 2 ? headings : [], '旧版 .doc 没有标题信息，未能生成目录');
  const zoom = cssZoom(scroller, body, saved.zoom ?? 1, tracker.refresh);
  scroller.scrollTop = (saved.scroll ?? 0) * Math.max(0, scroller.scrollHeight - scroller.clientHeight);

  return {
    scroller,
    zoom,
    state: () => ({
      zoom: zoom.get(),
      scroll: scroller.scrollTop / Math.max(1, scroller.scrollHeight - scroller.clientHeight),
    }),
    destroy: tracker.destroy,
  };
};
