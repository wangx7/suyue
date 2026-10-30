import DOMPurify from 'dompurify';
import { openLightbox } from '../lightbox';
import { el } from '../ui';
import { cssZoom } from '../zoom';
import { trackHeadings, type Heading } from './outline';
import type { Renderer } from './types';
import { find, kid, kids, Pkg, resolvePath } from './zip';

// 电子书：按 spine 顺序把各章 XHTML 接成一篇连续滚动的长文，排版沿用 Markdown 的样式，不用书自带的 CSS
// （深色主题、字号都由素阅控制）。目录取 EPUB3 的 nav 或 EPUB2 的 toc.ncx。

const ABSOLUTE = /^(?:[a-z][a-z\d+.-]*:|\/\/)/i;
const FONT = /\.(ttf|otf|woff2?)$/i;

interface TocEntry {
  title: string;
  href: string;
  level: number;
}

const clean = (s: string | null | undefined) => (s ?? '').replace(/\s+/g, ' ').trim();

/** EPUB3 导航文档：nav[epub:type=toc] 里嵌套的 ol/li */
function navToc(doc: Document, base: string): TocEntry[] {
  const navs = [...doc.getElementsByTagNameNS('*', 'nav')];
  const nav = navs.find((n) => [...n.attributes].some((a) => a.localName === 'type' && a.value.split(/\s+/).includes('toc'))) ?? navs[0];
  const out: TocEntry[] = [];
  const walk = (ol: Element | null, level: number) => {
    for (const li of kids(ol, 'li')) {
      const a = kid(li, 'a') ?? kid(li, 'span');
      const href = a?.getAttribute('href');
      const title = clean(a?.textContent);
      if (title) out.push({ title, href: href ? resolvePath(base, href) + (href.includes('#') ? href.slice(href.indexOf('#')) : '') : '', level });
      walk(kid(li, 'ol'), level + 1);
    }
  };
  walk(kid(nav, 'ol'), 1);
  return out;
}

/** EPUB2：toc.ncx 的 navPoint 树 */
function ncxToc(doc: Document, base: string): TocEntry[] {
  const out: TocEntry[] = [];
  const walk = (parent: Element | null, level: number) => {
    for (const p of kids(parent, 'navPoint')) {
      const title = clean(find(kid(p, 'navLabel'), 'text')?.textContent);
      const src = kid(p, 'content')?.getAttribute('src') ?? '';
      if (title) out.push({ title, href: src ? resolvePath(base, src) + (src.includes('#') ? src.slice(src.indexOf('#')) : '') : '', level });
      walk(p, level + 1);
    }
  };
  walk(find(doc, 'navMap'), 1);
  return out;
}

export const renderEpub: Renderer = async (data, { host, shell, saved }) => {
  const pkg = await Pkg.open(data);
  const opfPath = find(await pkg.xml('META-INF/container.xml'), 'rootfile')?.getAttribute('full-path');
  const opf = opfPath ? await pkg.xml(opfPath) : null;
  if (!opfPath || !opf) throw new Error('不是有效的 EPUB 文件');

  // 只混淆了字体的不算加密；正文被加密说明有 DRM
  const enc = await pkg.xml('META-INF/encryption.xml');
  const encrypted = [...(enc?.getElementsByTagNameNS('*', 'CipherReference') ?? [])].map((c) => c.getAttribute('URI') ?? '');
  if (encrypted.some((u) => !FONT.test(u))) throw new Error('这本电子书受 DRM 保护，无法打开');

  const manifest = new Map<string, { path: string; type: string; props: string }>();
  for (const item of kids(find(opf, 'manifest'), 'item')) {
    const id = item.getAttribute('id');
    const href = item.getAttribute('href');
    if (id && href) manifest.set(id, { path: resolvePath(opfPath, href), type: item.getAttribute('media-type') ?? '', props: item.getAttribute('properties') ?? '' });
  }
  const spineEl = find(opf, 'spine');
  const spine = kids(spineEl, 'itemref')
    .map((r) => manifest.get(r.getAttribute('idref') ?? ''))
    .filter((m): m is { path: string; type: string; props: string } => !!m && /html|xml/i.test(m.type));
  if (!spine.length) throw new Error('这本电子书没有正文');

  const scroller = el('div', 'doc-scroll');
  const body = el('article', 'md-read epub-read');
  scroller.append(body);
  host.append(scroller);

  // 每章的路径 → 章节序号，书内链接和目录都靠它定位
  const chapterOf = new Map(spine.map((s, i) => [s.path.toLowerCase(), i]));
  const target = (href: string): string | null => {
    const [path, frag] = href.split('#');
    const i = chapterOf.get(path.toLowerCase());
    if (i == null) return null;
    return frag ? `c${i}-${frag}` : `c${i}`;
  };

  // 一起读：JSZip 每次读取都要排队等一个定时器，逐章读的话上千章的书要好几秒
  const docs = await Promise.all(spine.map(async (ch) => (await pkg.xml(ch.path, 'application/xhtml+xml')) ?? pkg.xml(ch.path, 'text/html')));
  for (const [i, ch] of spine.entries()) {
    const doc = docs[i];
    const src = doc?.getElementsByTagNameNS('*', 'body')[0];
    const section = el('section', 'epub-chapter');
    section.id = `c${i}`;
    if (!src) {
      body.append(section);
      continue;
    }
    const html = DOMPurify.sanitize(src.innerHTML, {
      FORBID_TAGS: ['style', 'link', 'form', 'input', 'button', 'textarea', 'select'],
      FORBID_ATTR: ['style'],
      ADD_ATTR: ['xlink:href'],
    });
    const tpl = el('template');
    tpl.innerHTML = html;
    const content = tpl.content;

    // id 加上章节前缀，避免不同章节的锚点重名
    for (const n of content.querySelectorAll<HTMLElement>('[id]')) n.id = `c${i}-${n.id}`;
    for (const a of content.querySelectorAll<HTMLAnchorElement>('a[href]')) {
      const href = a.getAttribute('href')!;
      if (ABSOLUTE.test(href)) {
        a.target = '_blank';
        a.rel = 'noopener noreferrer';
        continue;
      }
      const t = href.startsWith('#') ? `c${i}-${href.slice(1)}` : target(resolvePath(ch.path, href) + (href.includes('#') ? href.slice(href.indexOf('#')) : ''));
      if (t) a.dataset.target = t;
      a.setAttribute('href', '#');
    }
    // 图片先记下包内路径，滚到附近再解压
    for (const img of content.querySelectorAll('img')) {
      const s = img.getAttribute('src');
      img.removeAttribute('src');
      img.removeAttribute('srcset');
      if (s && !ABSOLUTE.test(s)) img.dataset.src = resolvePath(ch.path, s);
      img.decoding = 'async';
    }
    for (const image of content.querySelectorAll('image')) {
      const s = image.getAttribute('href') ?? image.getAttribute('xlink:href');
      if (s && !ABSOLUTE.test(s)) image.setAttribute('data-src', resolvePath(ch.path, s));
      image.removeAttribute('href');
      image.removeAttribute('xlink:href');
    }
    section.append(content);
    body.append(section);
  }

  const io = new IntersectionObserver(
    (entries) => {
      for (const e of entries) {
        if (!e.isIntersecting) continue;
        io.unobserve(e.target);
        const node = e.target as HTMLElement | SVGElement;
        void pkg.url(node.dataset.src!).then((url) => {
          if (!url) return;
          if (node instanceof HTMLImageElement) node.src = url;
          else node.setAttribute('href', url);
        });
      }
    },
    { root: scroller, rootMargin: '200% 0px' },
  );
  body.querySelectorAll<HTMLElement>('[data-src]').forEach((n) => io.observe(n));

  // 目录：优先用书自带的；没有就用正文里的标题
  const navItem = [...manifest.values()].find((m) => m.props.split(/\s+/).includes('nav'));
  const ncxItem = manifest.get(spineEl?.getAttribute('toc') ?? '') ?? [...manifest.values()].find((m) => m.type === 'application/x-dtbncx+xml');
  let toc: TocEntry[] = [];
  const navDoc = navItem && (await pkg.xml(navItem.path, 'application/xhtml+xml'));
  if (navItem && navDoc) toc = navToc(navDoc, navItem.path);
  if (!toc.length && ncxItem) {
    const ncx = await pkg.xml(ncxItem.path);
    if (ncx) toc = ncxToc(ncx, ncxItem.path);
  }
  const headings: Heading[] = [];
  for (const entry of toc) {
    const id = entry.href && target(entry.href);
    // 锚点找不到时退回到所在章节的开头
    const byId = (s: string) => body.querySelector<HTMLElement>(`[id="${CSS.escape(s)}"]`);
    const node = id ? (byId(id) ?? byId(id.replace(/-.*$/, ''))) : null;
    if (node) headings.push({ el: node, item: { id: `t${headings.length}`, title: entry.title, level: entry.level } });
  }
  if (!headings.length) {
    for (const h of body.querySelectorAll<HTMLElement>('h1, h2, h3')) {
      const title = clean(h.textContent);
      if (title) headings.push({ el: h, item: { id: `t${headings.length}`, title, level: Number(h.tagName[1]) } });
    }
  }

  const tracker = trackHeadings(scroller, shell);
  tracker.set(headings, '这本书没有目录');
  const zoom = cssZoom(scroller, body, saved.zoom ?? 1, tracker.refresh);

  body.addEventListener('click', (e) => {
    const t = e.target as Element;
    const a = t.closest?.<HTMLAnchorElement>('a[href]');
    if (!a) {
      const img = t.closest?.<HTMLImageElement>('img');
      if (img?.currentSrc) openLightbox(img);
      return;
    }
    if (a.target === '_blank') return;
    e.preventDefault();
    const id = a.dataset.target;
    const dest = id && body.querySelector(`[id="${CSS.escape(id)}"]`);
    if (dest) scroller.scrollTo({ top: scroller.scrollTop + dest.getBoundingClientRect().top - scroller.getBoundingClientRect().top - 24, behavior: 'smooth' });
  });

  scroller.scrollTop = (saved.scroll ?? 0) * Math.max(0, scroller.scrollHeight - scroller.clientHeight);

  return {
    scroller,
    zoom,
    state: () => ({
      zoom: zoom.get(),
      scroll: scroller.scrollTop / Math.max(1, scroller.scrollHeight - scroller.clientHeight),
    }),
    destroy() {
      tracker.destroy();
      io.disconnect();
      pkg.destroy();
    },
  };
};
