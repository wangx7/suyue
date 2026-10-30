// VS Code 自定义编辑器里的素阅：扩展读出文件发过来，渲染部分和网页版共用
import '../../src/styles.css';
import './vscode.css';
import { detect, KIND_LABEL, loaders, sizeText } from '../../src/formats';
import type { Linker, Viewer } from '../../src/renderers/types';
import { createShell, scopedShell } from '../../src/shell';
import type { ReadState } from '../../src/store';
import { initTheme, onThemeChange, setHostDark, THEME_LABEL } from '../../src/theme';
import { el } from '../../src/ui';
import { attachZoom } from '../../src/zoom';
import { renderAudio } from './media';

declare function acquireVsCodeApi(): { postMessage(msg: unknown): void };

interface LoadMessage {
  type: 'load';
  name: string;
  /** 文档内容；音频不传内容，改传 src */
  data?: Uint8Array | ArrayBuffer;
  /** 音频的 webview 资源地址 */
  src?: string;
  saved: ReadState;
  /** 文档所在目录的 webview 地址，以 / 结尾 */
  base: string;
}

const vscode = acquireVsCodeApi();

// VS Code 默认给 webview 注入的样式（body 内边距、链接和代码颜色）会盖掉素阅的排版
document.getElementById('_defaultStyles')?.remove();

// 主题和侧栏开合由扩展保存：先把上次的值放进 localStorage，再把之后的改动转给扩展
try {
  const prefs = JSON.parse(document.querySelector<HTMLMetaElement>('meta[name="suyue-prefs"]')?.content || '{}') as Record<string, string>;
  for (const [k, v] of Object.entries(prefs)) localStorage.setItem(k, v);
  const setItem = Storage.prototype.setItem;
  Storage.prototype.setItem = function (this: Storage, key: string, value: string) {
    setItem.call(this, key, value);
    vscode.postMessage({ type: 'pref', key, value });
  };
} catch {
  /* localStorage 不可用时用默认设置 */
}

// “跟随系统”在这里是跟随 VS Code 的主题
THEME_LABEL.system = '跟随 VS Code';
initTheme();
const hostDark = () => {
  const c = document.body.classList;
  return c.contains('vscode-dark') || (c.contains('vscode-high-contrast') && !c.contains('vscode-high-contrast-light'));
};
setHostDark(hostDark());
new MutationObserver(() => setHostDark(hostDark())).observe(document.body, { attributes: true, attributeFilter: ['class'] });

const shell = createShell({
  onHome() {},
  onZoomReset() {
    if (!current?.zoom) return;
    current.zoom.set(1, { x: 0, y: 0 });
    onZoom(1);
  },
});
const stage = shell.stage;

let current: Viewer | null = null;
let token = 0;
let saveTimer = 0;

function saveNow() {
  clearTimeout(saveTimer);
  if (current?.state) vscode.postMessage({ type: 'state', state: current.state() });
}
function saveSoon() {
  clearTimeout(saveTimer);
  saveTimer = window.setTimeout(saveNow, 1000);
}

function onZoom(z: number) {
  shell.setZoom(z);
  shell.hud(`${Math.round(z * 100)}%`);
  saveSoon();
}

function message(title: string, detail: string) {
  const wrap = el('div', 'center');
  const box = el('div', 'message');
  box.append(el('h1', '', title), el('p', '', detail));
  wrap.append(box);
  stage.append(wrap);
}

/** 相对路径的图片按文档目录取；链接交给扩展，在 VS Code 里打开 */
function linker(base: string): Linker {
  return {
    async asset(path) {
      try {
        return new URL(path, base).href;
      } catch {
        return null;
      }
    },
    open(path) {
      vscode.postMessage({ type: 'open', path });
      return true;
    },
  };
}

async function open({ name, data, src, saved, base }: LoadMessage) {
  const t = ++token;
  // 文件改动后重新加载时沿用当前的阅读位置
  if (current) {
    saved = current.state?.() ?? saved;
    current.destroy();
    current = null;
  }
  stage.replaceChildren();
  shell.setActions([]);
  shell.setOutline(null);
  shell.setZoom(null);
  shell.watchScroll(null);

  if (src) {
    shell.setDoc(name, '音频');
    const host = el('div', 'viewer-host');
    stage.append(host);
    current = renderAudio(src, host, name, saved, saveSoon);
    return;
  }
  if (!data) return;
  const buf = ArrayBuffer.isView(data) ? (data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength) as ArrayBuffer) : data;
  const kind = detect(name, buf);
  if (!kind) {
    shell.setDoc(null);
    return message('无法识别这个文件', `「${name}」不是素阅支持的格式。`);
  }
  shell.setDoc(name, `${KIND_LABEL[kind]} · ${sizeText(buf.byteLength)}`);
  const host = el('div', 'viewer-host');
  const loading = el('div', 'center loading');
  loading.append(el('div', 'spinner'), el('div', '', '正在打开…'));
  stage.append(host, loading);
  try {
    const render = await loaders[kind]();
    const viewer = await render(buf, { host, shell: scopedShell(shell, () => t === token), name, saved, linker: linker(base) });
    if (t !== token) return viewer.destroy();
    current = viewer;
    loading.remove();
    shell.setZoom(viewer.zoom?.get() ?? null);
    shell.watchScroll(viewer.scroller ?? null);
    viewer.scroller?.addEventListener('scroll', saveSoon, { passive: true });
  } catch (err) {
    if (t !== token) return;
    console.error(err);
    stage.replaceChildren();
    message('打开失败', err instanceof Error ? err.message : String(err));
  }
}

onThemeChange((dark) => current?.onTheme?.(dark));
attachZoom(stage, () => current?.zoom, onZoom);

window.addEventListener('message', (e: MessageEvent<LoadMessage>) => {
  if (e.data?.type === 'load') open(e.data);
});
// 表格接管 Ctrl+F；Word 和 Markdown 用 VS Code 的查找框
window.addEventListener('keydown', (e) => {
  if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'f' && current?.openSearch) {
    e.preventDefault();
    current.openSearch();
  }
});
window.addEventListener('pagehide', saveNow);
document.addEventListener('visibilitychange', () => document.visibilityState === 'hidden' && saveNow());

vscode.postMessage({ type: 'ready' });
