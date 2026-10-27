import DOMPurify from 'dompurify';
import hljs from 'highlight.js/lib/common';
import MarkdownIt from 'markdown-it';
import { icon } from '../icons';
import { openLightbox } from '../lightbox';
import { el } from '../ui';
import { cssZoom } from '../zoom';
import { trackHeadings, type Heading } from './outline';
import type { Renderer } from './types';

// Markdown：markdown-it 解析（CommonMark + GFM 表格、删除线、自动链接），允许内嵌 HTML，
// 但经 DOMPurify 过滤（文件可能来自网络，查看器页面在扩展里有较高权限）。配色全部走主题变量。

const md = new MarkdownIt({
  html: true,
  linkify: true,
  highlight(code, lang) {
    if (!lang || !hljs.getLanguage(lang)) return '';
    try {
      return hljs.highlight(code, { language: lang, ignoreIllegals: true }).value;
    } catch {
      return '';
    }
  },
});

/** 绝对地址：带协议（http:、data:、blob: 等）或 // 开头 */
const ABSOLUTE = /^(?:[a-z][a-z\d+.-]*:|\/\/)/i;

/** 先按 UTF-8 严格解码，失败再按 GBK（国内 Windows 上的老文件） */
function decode(buf: ArrayBuffer): string {
  const b = new Uint8Array(buf);
  if (b[0] === 0xff && b[1] === 0xfe) return new TextDecoder('utf-16le').decode(b);
  if (b[0] === 0xfe && b[1] === 0xff) return new TextDecoder('utf-16be').decode(b);
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(b);
  } catch {
    return new TextDecoder('gbk').decode(b);
  }
}

/** 去掉开头的 YAML front matter */
const stripFrontMatter = (src: string) => src.replace(/^---\r?\n[\s\S]*?\r?\n(?:---|\.\.\.)[ \t]*(?:\r?\n|$)/, '');

/** 与 GitHub 一致的标题锚点：小写，去掉标点，空白换成 -，重名加序号 */
function slugger() {
  const seen = new Map<string, number>();
  return (text: string) => {
    const base = text.trim().toLowerCase().replace(/[^\p{L}\p{N}\s_-]/gu, '').replace(/\s/g, '-');
    const n = seen.get(base) ?? 0;
    seen.set(base, n + 1);
    return n ? `${base}-${n}` : base;
  };
}

/** “- [ ] 待办” / “- [x] 已完成” */
function taskLists(root: ParentNode) {
  for (const li of root.querySelectorAll('li')) {
    const host = li.firstElementChild?.tagName === 'P' && li.firstChild === li.firstElementChild ? li.firstElementChild : li;
    const text = host.firstChild;
    const m = text?.nodeType === Node.TEXT_NODE ? /^\[([ xX])\]\s+/.exec(text.textContent!) : null;
    if (!m) continue;
    text!.textContent = text!.textContent!.slice(m[0].length);
    const box = el('input');
    box.type = 'checkbox';
    box.disabled = true;
    box.checked = m[1] !== ' ';
    host.prepend(box);
    li.classList.add('task-list-item');
    li.parentElement?.classList.add('contains-task-list');
  }
}

/** 代码块：悬停出现复制按钮 */
function codeBlocks(root: ParentNode) {
  for (const pre of root.querySelectorAll('pre')) {
    const code = pre.querySelector(':scope > code');
    if (!code) continue;
    const wrap = el('div', 'md-code');
    pre.replaceWith(wrap);
    const copy = el('button', 'md-code-copy');
    copy.title = '复制代码';
    copy.append(icon('copy'));
    copy.addEventListener('click', async () => {
      try {
        await navigator.clipboard.writeText(code.textContent ?? '');
        copy.replaceChildren(icon('check'));
        window.setTimeout(() => copy.replaceChildren(icon('copy')), 1200);
      } catch {
        /* 剪贴板不可用 */
      }
    });
    wrap.append(copy, pre);
  }
}

export const renderMarkdown: Renderer = async (data, { host, shell, saved, linker }) => {
  const html = DOMPurify.sanitize(md.render(stripFrontMatter(decode(data))), {
    FORBID_TAGS: ['style', 'form', 'input', 'button', 'textarea', 'select'],
  });
  // 先放进 template（惰性文档），相对路径的图片在改写前不会发出请求
  const tpl = el('template');
  tpl.innerHTML = html;
  const content = tpl.content;

  const scroller = el('div', 'doc-scroll');
  const body = el('article', 'md-read');
  scroller.append(body);
  host.append(scroller);

  taskLists(content);
  codeBlocks(content);
  for (const t of content.querySelectorAll('table')) {
    const wrap = el('div', 'md-table');
    t.replaceWith(wrap);
    wrap.append(t);
  }
  for (const a of content.querySelectorAll<HTMLAnchorElement>('a[href]')) {
    if (ABSOLUTE.test(a.getAttribute('href')!)) {
      a.target = '_blank';
      a.rel = 'noopener noreferrer';
    }
  }

  // 相对路径的图片交给 linker 按文档网址解析；本地文件读不到同目录的图片，显示占位
  const missing = (img: HTMLImageElement) => {
    const ph = el('span', 'md-missing', `图片：${img.alt || img.dataset.src}`);
    ph.title = img.dataset.src ?? '';
    img.replaceWith(ph);
  };
  for (const img of content.querySelectorAll('img')) {
    img.loading = 'lazy';
    img.decoding = 'async';
    const src = img.getAttribute('src') ?? '';
    if (!src || ABSOLUTE.test(src)) continue;
    img.removeAttribute('src');
    img.dataset.src = src;
    if (!linker) {
      missing(img);
      continue;
    }
    linker.asset(src).then(
      (url) => (url ? (img.src = url) : missing(img)),
      () => missing(img),
    );
  }

  const slug = slugger();
  const headings: Heading[] = [];
  for (const h of content.querySelectorAll<HTMLHeadingElement>('h1, h2, h3, h4, h5, h6')) {
    const title = h.textContent?.replace(/\s+/g, ' ').trim();
    if (!title) continue;
    h.id ||= slug(title);
    headings.push({ el: h, item: { id: `h${headings.length}`, title, level: Number(h.tagName[1]) } });
  }
  body.append(content);

  const tracker = trackHeadings(scroller, shell);
  tracker.set(headings, '此文档没有标题');
  const zoom = cssZoom(scroller, body, saved.zoom ?? 1, tracker.refresh);

  body.addEventListener('click', (e) => {
    const target = e.target as Element;
    const a = target.closest?.<HTMLAnchorElement>('a[href]');
    if (!a) {
      const img = target.closest?.<HTMLImageElement>('img');
      if (img?.currentSrc) openLightbox(img);
      return;
    }
    const href = a.getAttribute('href')!;
    if (ABSOLUTE.test(href)) return;
    e.preventDefault();
    if (href.startsWith('#')) {
      let id = href.slice(1);
      try {
        id = decodeURIComponent(id);
      } catch {
        /* 保持原样 */
      }
      const dest = id ? body.querySelector(`[id="${CSS.escape(id)}"], [name="${CSS.escape(id)}"]`) : null;
      if (dest) scroller.scrollTo({ top: scroller.scrollTop + dest.getBoundingClientRect().top - scroller.getBoundingClientRect().top - 24, behavior: 'smooth' });
      return;
    }
    if (!linker?.open(href)) shell.hud('本地文件无法跳转到同目录的其他文件', 1600);
  });

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
