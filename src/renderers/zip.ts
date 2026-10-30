import JSZip from 'jszip';

// PPTX 和 EPUB 共用：zip 包里按路径读 XML、读图片（转成 blob 地址），以及 XML 的小工具。
// XML 元素一律按 localName 查找，不管命名空间前缀（不同软件写出的前缀不一定相同）。

const MIME: Record<string, string> = {
  png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', jpe: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp',
  webp: 'image/webp', avif: 'image/avif', svg: 'image/svg+xml', ico: 'image/x-icon',
};

/** 把 target 按 base 所在目录解析成包内路径（不带开头的 /） */
export function resolvePath(base: string, target: string): string {
  let t = target.split('#')[0];
  try {
    t = decodeURIComponent(t);
  } catch {
    /* 保持原样 */
  }
  const parts = t.startsWith('/') ? [] : base.split('/').slice(0, -1);
  for (const seg of t.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg && seg !== '.') parts.push(seg);
  }
  return parts.join('/');
}

export class Pkg {
  private files = new Map<string, JSZip.JSZipObject>();
  private docs = new Map<string, Promise<Document | null>>();
  private urls = new Map<string, Promise<string | null>>();

  private constructor(zip: JSZip) {
    // 有的软件写出的路径大小写和引用处不一致，统一按小写找
    zip.forEach((path, file) => !file.dir && this.files.set(path.toLowerCase(), file));
  }

  static async open(data: ArrayBuffer): Promise<Pkg> {
    return new Pkg(await JSZip.loadAsync(data));
  }

  has(path: string): boolean {
    return this.files.has(path.toLowerCase());
  }

  /** JSZip 自带的 UTF-8 解码是纯 JS 实现，大文件很慢，取字节后交给 TextDecoder */
  async text(path: string): Promise<string | null> {
    const bytes = await this.files.get(path.toLowerCase())?.async('uint8array');
    return bytes ? new TextDecoder().decode(bytes) : null;
  }

  /** 解析失败或文件不存在时返回 null；同一路径只解析一次 */
  xml(path: string, type: DOMParserSupportedType = 'application/xml'): Promise<Document | null> {
    const key = `${type}:${path.toLowerCase()}`;
    let p = this.docs.get(key);
    if (!p) {
      p = this.text(path).then((s) => {
        if (s == null) return null;
        const doc = new DOMParser().parseFromString(s, type);
        return doc.getElementsByTagName('parsererror').length ? null : doc;
      });
      this.docs.set(key, p);
    }
    return p;
  }

  /** 包内图片的 blob 地址；浏览器显示不了的格式（emf、wmf、tiff）返回 null */
  url(path: string): Promise<string | null> {
    const key = path.toLowerCase();
    let p = this.urls.get(key);
    if (!p) {
      const type = MIME[/\.([a-z0-9]+)$/i.exec(key)?.[1] ?? ''];
      const file = this.files.get(key);
      p = type && file ? file.async('arraybuffer').then((b) => URL.createObjectURL(new Blob([b], { type }))) : Promise.resolve(null);
      this.urls.set(key, p);
    }
    return p;
  }

  destroy() {
    for (const p of this.urls.values()) p.then((u) => u && URL.revokeObjectURL(u));
    this.urls.clear();
  }
}

export function kid(node: Element | null | undefined, name: string): Element | null {
  if (!node) return null;
  for (let c = node.firstElementChild; c; c = c.nextElementSibling) if (c.localName === name) return c;
  return null;
}

export function kids(node: Element | null | undefined, name?: string): Element[] {
  const out: Element[] = [];
  if (!node) return out;
  for (let c = node.firstElementChild; c; c = c.nextElementSibling) if (!name || c.localName === name) out.push(c);
  return out;
}

/** 逐级取子元素：dig(sp, 'spPr', 'xfrm') */
export function dig(node: Element | null | undefined, ...names: string[]): Element | null {
  let cur = node ?? null;
  for (const n of names) cur = kid(cur, n);
  return cur;
}

/** 文档里第一个 localName 匹配的元素（任意深度） */
export function find(root: Document | Element | null | undefined, name: string): Element | null {
  if (!root) return null;
  return root.getElementsByTagNameNS('*', name)[0] ?? null;
}

export function num(node: Element | null | undefined, attr: string): number | undefined;
export function num(node: Element | null | undefined, attr: string, def: number): number;
export function num(node: Element | null | undefined, attr: string, def?: number): number | undefined {
  const v = node?.getAttribute(attr);
  if (v == null || v === '') return def;
  const n = Number(v);
  return Number.isFinite(n) ? n : def;
}

/** 带命名空间前缀的属性（r:embed、r:id 等），按 localName 找 */
export function nsAttr(node: Element | null | undefined, name: string): string | null {
  if (!node) return null;
  for (const a of node.attributes) if (a.localName === name && a.prefix) return a.value;
  return null;
}
