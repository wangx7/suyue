import './styles.css';
import { createShell, recentList, scopedShell } from './shell';
import { addRecent, fileKey, listRecent, loadState, removeRecent, saveState, urlKey, type RecentItem } from './store';
import { initTheme, onThemeChange } from './theme';
import { button, el } from './ui';
import { attachZoom } from './zoom';
import { ALL_EXTS, detect, extOf, KIND_LABEL, loaders, MD_EXTS, sizeText } from './formats';
import type { Linker, Viewer } from './renderers/types';

const APP_NAME = '素阅';

// 浏览器自己能显示的格式，交给新标签页；xml、json 按纯文本显示
const TEXT = 'text/plain;charset=utf-8';
const BROWSER_TYPES: Record<string, string> = {
  pdf: 'application/pdf',
  png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', webp: 'image/webp', avif: 'image/avif', bmp: 'image/bmp', ico: 'image/x-icon',
  mp4: 'video/mp4', m4v: 'video/mp4', webm: 'video/webm', ogv: 'video/ogg',
  mp3: 'audio/mpeg', m4a: 'audio/mp4', aac: 'audio/aac', wav: 'audio/wav', ogg: 'audio/ogg', oga: 'audio/ogg', opus: 'audio/ogg', flac: 'audio/flac',
  json: TEXT, xml: TEXT, log: TEXT, yaml: TEXT, yml: TEXT,
  html: 'text/html', htm: 'text/html', svg: 'image/svg+xml',
};
const ACCEPT: Record<string, string[]> = {
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx', '.docm'],
  'application/msword': ['.doc'],
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx', '.xlsm'],
  'application/vnd.ms-excel.sheet.binary.macroEnabled.12': ['.xlsb'],
  'application/vnd.ms-excel': ['.xls'],
  'application/vnd.oasis.opendocument.spreadsheet': ['.ods'],
  'text/csv': ['.csv'],
  'text/tab-separated-values': ['.tsv'],
  'text/markdown': MD_EXTS.map((e) => `.${e}`),
  'application/vnd.openxmlformats-officedocument.presentationml.presentation': ['.pptx', '.pptm'],
  'application/vnd.openxmlformats-officedocument.presentationml.slideshow': ['.ppsx', '.ppsm'],
  'application/vnd.openxmlformats-officedocument.presentationml.template': ['.potx'],
  'application/epub+zip': ['.epub'],
};
const isExtension = typeof chrome !== 'undefined' && !!chrome.runtime?.id;

initTheme();
const shell = createShell({
  onHome: showHome,
  onZoomReset() {
    if (!current?.zoom) return;
    current.zoom.set(1, { x: 0, y: 0 });
    onZoom(1);
  },
});
const stage = shell.stage;

const fileInput = el('input');
fileInput.type = 'file';
fileInput.hidden = true;
fileInput.accept = Object.values(ACCEPT).flat().join(',');
const dropOverlay = el('div', 'drop-overlay', '松开以打开');
document.body.append(fileInput, dropOverlay);

let current: Viewer | null = null;
let currentKey = '';
let openToken = 0;

/** 浏览器能直接显示时返回 MIME 类型：先看扩展名，没有再看文件头 */
function browserType(name: string, buf: ArrayBuffer): string | null {
  const known = BROWSER_TYPES[extOf(name)];
  if (known) return known;
  const b = new Uint8Array(buf, 0, Math.min(12, buf.byteLength));
  const head = String.fromCharCode(...b);
  if (head.startsWith('%PDF')) return 'application/pdf';
  if (b[0] === 0x89 && head.slice(1, 4) === 'PNG') return 'image/png';
  if (b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff) return 'image/jpeg';
  if (head.startsWith('GIF8')) return 'image/gif';
  if (head.startsWith('RIFF') && head.slice(8, 12) === 'WEBP') return 'image/webp';
  return null;
}

/**
 * 本地文件的 blob 地址。blob 地址随本页关闭失效，所以不主动回收，新标签页里刷新、另存都还能用。
 * blob 和查看器同源，网页、svg 里的脚本能读到本页的存储，所以套一层不给同源权限的沙箱 iframe
 */
function blobHref(buf: ArrayBuffer, type: string, name: string): string {
  if (type !== 'text/html' && type !== 'image/svg+xml') return URL.createObjectURL(new Blob([buf], { type }));
  // 没写 meta charset 的网页，浏览器会按系统语言猜编码；能按 UTF-8 解开就明确标上
  let charset = '';
  try {
    new TextDecoder('utf-8', { fatal: true }).decode(buf);
    charset = ';charset=utf-8';
  } catch {
    /* 交给网页自己的 meta charset */
  }
  const inner = URL.createObjectURL(new Blob([buf], { type: type + charset }));
  const title = name.replace(/[&<>]/g, (c) => `&#${c.charCodeAt(0)};`);
  const page =
    `<!doctype html><meta charset="utf-8"><title>${title}</title>` +
    '<style>html,body,iframe{margin:0;width:100%;height:100%;border:0;display:block}</style>' +
    `<iframe sandbox="allow-scripts allow-forms allow-popups allow-modals allow-downloads" src="${inner}"></iframe>`;
  return URL.createObjectURL(new Blob([page], { type: 'text/html;charset=utf-8' }));
}

// ---------- 阅读位置 ----------
function saveCurrent() {
  if (current?.state && currentKey) saveState(currentKey, current.state());
}
let saveTimer = 0;
function saveSoon() {
  clearTimeout(saveTimer);
  saveTimer = window.setTimeout(saveCurrent, 1000);
}

function onZoom(z: number) {
  shell.setZoom(z);
  shell.hud(`${Math.round(z * 100)}%`);
  saveSoon();
}

// ---------- 正文区状态 ----------
function clearStage() {
  saveCurrent();
  current?.destroy();
  current = null;
  currentKey = '';
  stage.replaceChildren();
  shell.setActions([]);
  shell.setOutline(null);
  shell.setZoom(null);
  shell.watchScroll(null);
}

function center(...children: Node[]) {
  const wrap = el('div', 'center');
  wrap.append(...children);
  stage.append(wrap);
  return wrap;
}

/** 首页：名字、打开文件、最近打开 */
async function showHome() {
  const token = ++openToken;
  clearStage();
  shell.setDoc(null);
  document.title = APP_NAME;
  const home = el('div', 'home');
  home.append(
    el('h1', 'home-name', APP_NAME),
    button('选择文件', '⌘/Ctrl + O', pickFile),
    el('p', 'home-hint', '或拖入 Word、Excel、PPT、电子书、CSV、Markdown 文件'),
  );
  center(home, el('p', 'home-foot', '文件只在本地解析，不会上传'));
  const items = (await listRecent()).slice(0, 8);
  if (token !== openToken || !items.length) return;
  const recent = el('section', 'home-recent');
  recent.append(
    el('h2', '', '最近'),
    recentList(items, {
      onOpen: openRecent,
      async onRemove(item) {
        await removeRecent(item.id);
        showHome();
      },
    }),
  );
  home.append(recent);
}

function showLoading(text: string): HTMLElement {
  const box = el('div', 'center loading');
  box.append(el('div', 'spinner'), el('div', '', text));
  stage.append(box);
  return box;
}

function showError(title: string, detail: string, action?: { label: string; run: () => void }) {
  clearStage();
  shell.setDoc(null);
  document.title = APP_NAME;
  const box = el('div', 'message');
  const { label, run } = action ?? { label: '返回首页', run: showHome };
  box.append(el('h1', '', title), el('p', '', detail), button(label, '', run));
  center(box);
}

/** 在新标签页打开；网页版不在点击里时可能被拦截，就留一个按钮让用户点 */
function openInTab(href: string, name: string) {
  const open = () => {
    if (isExtension) {
      chrome.tabs.create({ url: href });
      return true;
    }
    const w = window.open(href, '_blank');
    if (w) w.opener = null;
    return !!w;
  };
  if (!open()) {
    return showError('在新标签页打开', `「${name}」由浏览器直接显示。`, {
      label: '在新标签页打开',
      run() {
        if (open()) showHome();
      },
    });
  }
  showHome();
  shell.hud(`已在新标签页打开「${name}」`, 1600);
}

interface OpenOptions {
  /** 记入“最近打开”的额外字段（句柄、网址） */
  recent?: Partial<RecentItem>;
  /** 从“最近打开”重新打开时，旧记录的 id */
  replaces?: string;
  linker?: Linker;
}

async function openBuffer(buf: ArrayBuffer, name: string, key: string, { recent = {}, replaces, linker }: OpenOptions = {}) {
  const token = ++openToken;
  clearStage();
  const kind = detect(name, buf);
  const type = kind ? null : browserType(name, buf);
  if (type) return openInTab(recent.url ?? blobHref(buf, type, name), name);
  if (!kind) return showError('暂不支持这个文件', `无法识别「${name}」的格式。支持 Word、Excel、PPT、EPUB 电子书、CSV 和 Markdown。`);

  const size = buf.byteLength; // 表格解析时 buf 会转交给 Worker，先记下大小
  shell.setDoc(name, `${KIND_LABEL[kind]} · ${sizeText(size)}`);
  document.title = `${name} · ${APP_NAME}`;
  const host = el('div', 'viewer-host');
  stage.append(host);
  const loading = showLoading('正在打开…');
  try {
    const render = await loaders[kind]();
    const viewer = await render(buf, { host, shell: scopedShell(shell, () => token === openToken), name, saved: loadState(key), linker });
    if (token !== openToken) return viewer.destroy();
    current = viewer;
    currentKey = key;
    loading.remove();
    shell.setZoom(viewer.zoom?.get() ?? null);
    shell.watchScroll(viewer.scroller ?? null);
    viewer.scroller?.addEventListener('scroll', saveSoon, { passive: true });
    await addRecent({ ...recent, id: key, name, kind, size, time: Date.now() });
    if (replaces && replaces !== key) await removeRecent(replaces);
  } catch (err) {
    if (token !== openToken) return;
    console.error(err);
    showError('打开失败', err instanceof Error ? err.message : String(err));
  }
}

async function openFile(file: File, handle?: FileSystemFileHandle, replaces?: string) {
  try {
    await openBuffer(await file.arrayBuffer(), file.name, fileKey(file), { recent: handle ? { handle } : {}, replaces });
  } catch (err) {
    showError('读取文件失败', String(err));
  }
}

/** 优先用系统文件选择器拿到句柄，这样“最近打开”可以直接再次打开 */
async function pickFile() {
  if (!window.showOpenFilePicker) return fileInput.click();
  try {
    const [handle] = await window.showOpenFilePicker({ types: [{ description: 'Word / Excel / PPT / EPUB / CSV / Markdown', accept: ACCEPT }] });
    await openFile(await handle.getFile(), handle);
  } catch (err) {
    if ((err as DOMException).name !== 'AbortError') fileInput.click();
  }
}

async function openRecent(item: RecentItem) {
  if (item.url) return openUrl(item.url);
  const h = item.handle;
  if (!h) return shell.hud('这个文件需要重新选择', 1600);
  try {
    const granted =
      (await h.queryPermission?.({ mode: 'read' })) === 'granted' || (await h.requestPermission?.({ mode: 'read' })) === 'granted';
    if (!granted) return shell.hud('没有获得读取权限', 1600);
    await openFile(await h.getFile(), h, item.id);
  } catch {
    shell.hud('文件已移动或删除', 1600);
    await removeRecent(item.id);
    showHome();
  }
}

function nameFromUrl(url: string, disposition: string | null): string {
  const star = disposition && /filename\*\s*=\s*(?:UTF-8'')?([^;]+)/i.exec(disposition);
  const plain = disposition && /filename\s*=\s*"?([^";]+)"?/i.exec(disposition);
  const raw = star?.[1] ?? plain?.[1];
  if (raw) {
    try {
      return decodeURIComponent(raw.trim());
    } catch {
      return raw.trim();
    }
  }
  try {
    const seg = new URL(url, location.href).pathname.split('/').filter(Boolean).pop();
    return seg ? decodeURIComponent(seg) : 'document';
  } catch {
    return 'document';
  }
}

// fetch 不支持 file:// 协议，扩展里读本地文件链接要走 XHR
function readFileUrl(url: string): Promise<ArrayBuffer> {
  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    xhr.open('GET', url);
    xhr.responseType = 'arraybuffer';
    xhr.onload = () => (xhr.response ? resolve(xhr.response) : reject(new Error('空响应')));
    xhr.onerror = () => reject(new Error('无法读取本地文件'));
    xhr.send();
  });
}

/** 网址打开的文档：相对路径按文档网址解析，能打开的文档在查看器里打开，其他链接在新标签页打开 */
function urlLinker(base: string): Linker {
  const abs = (path: string) => {
    try {
      return new URL(path, base).href;
    } catch {
      return null;
    }
  };
  return {
    asset: async (path) => abs(path),
    open(path) {
      const url = abs(path);
      if (!url) return false;
      const ext = extOf(nameFromUrl(url, null));
      if (ALL_EXTS.includes(ext)) openUrl(url);
      else window.open(url, '_blank', 'noopener');
      return true;
    },
  };
}

async function openUrl(rawUrl: string) {
  let url: string;
  try {
    url = new URL(rawUrl, location.href).href;
  } catch {
    return showError('链接无效', rawUrl);
  }
  // 从扩展名就能看出是 PDF、图片这类，不必先下载
  const urlName = nameFromUrl(url, null);
  if (extOf(urlName) in BROWSER_TYPES) return openInTab(url, urlName);
  clearStage();
  const loading = showLoading('正在下载…');
  const isFile = url.startsWith('file:');
  let buf: ArrayBuffer;
  let name: string;
  try {
    if (isFile) {
      buf = await readFileUrl(url);
      name = nameFromUrl(url, null);
    } else {
      const res = await fetch(url, { credentials: isExtension ? 'include' : 'same-origin' });
      if (!res.ok) throw new Error(`服务器返回 ${res.status}`);
      buf = await res.arrayBuffer();
      name = nameFromUrl(url, res.headers.get('content-disposition'));
    }
  } catch (err) {
    loading.remove();
    if (isFile) {
      return showError('无法读取本地文件', isExtension
        ? `请到 chrome://extensions → ${APP_NAME} → 详情，打开「允许访问文件网址」后重试。`
        : '网页版无法直接读取本地路径，请把文件拖进来。');
    }
    const { origin, host } = new URL(url);
    const canAsk = isExtension && !!chrome.permissions && origin.startsWith('http');
    return showError('下载失败', `${err instanceof Error ? err.message : err}：${url}`, canAsk
      ? {
          label: `允许读取 ${host}`,
          run: async () => {
            if (await chrome.permissions.request({ origins: [`${origin}/*`] })) openUrl(url);
          },
        }
      : undefined);
  }
  await openBuffer(buf, name, urlKey(url), { recent: { url }, linker: urlLinker(url) });
}

// ---------- 事件 ----------
onThemeChange((dark) => current?.onTheme?.(dark));
attachZoom(stage, () => current?.zoom, onZoom);

fileInput.addEventListener('change', () => {
  const file = fileInput.files?.[0];
  if (file) openFile(file);
  fileInput.value = '';
});

let dragDepth = 0;
window.addEventListener('dragenter', (e) => {
  if (!e.dataTransfer?.types.includes('Files')) return;
  dragDepth++;
  dropOverlay.classList.add('show');
});
window.addEventListener('dragleave', () => {
  dragDepth = Math.max(0, dragDepth - 1);
  if (!dragDepth) dropOverlay.classList.remove('show');
});
window.addEventListener('dragover', (e) => e.preventDefault());
window.addEventListener('drop', async (e) => {
  e.preventDefault();
  dragDepth = 0;
  dropOverlay.classList.remove('show');
  const file = e.dataTransfer?.files[0];
  // 句柄必须在 drop 事件内同步发起获取，之后 DataTransfer 就失效了
  const handlePromise = e.dataTransfer?.items[0]?.getAsFileSystemHandle?.();
  if (!file) return;
  let handle: FileSystemFileHandle | undefined;
  try {
    const h = await handlePromise;
    if (h?.kind === 'file') handle = h as FileSystemFileHandle;
    else if (h?.kind === 'directory') return shell.hud('请拖入单个文件', 1600);
  } catch {
    /* 不支持时只是记不住句柄 */
  }
  openFile(file, handle);
});

window.addEventListener('keydown', (e) => {
  const mod = e.metaKey || e.ctrlKey;
  const key = e.key.toLowerCase();
  if (mod && key === 'o') {
    e.preventDefault();
    pickFile();
  } else if (mod && e.key === '\\') {
    e.preventDefault();
    shell.toggleSide();
  } else if (mod && key === 'f' && current?.openSearch) {
    // 表格接管 Ctrl+F；Word 和 Markdown 走浏览器原生查找
    e.preventDefault();
    current.openSearch();
  }
});

window.addEventListener('pagehide', saveCurrent);
document.addEventListener('visibilitychange', () => document.visibilityState === 'hidden' && saveCurrent());

const urlParam = new URLSearchParams(location.search).get('url');
if (urlParam) openUrl(urlParam);
else showHome();
