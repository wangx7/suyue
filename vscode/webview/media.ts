// 音频：VS Code 自带的预览不认 flac，这里用 webview 里的 <audio> 播放。
// 文件不经过 postMessage，直接用 webview 资源地址，拖动进度时按需读取。
import { icon } from '../../src/icons';
import type { Viewer } from '../../src/renderers/types';
import type { ReadState } from '../../src/store';
import { el } from '../../src/ui';

const STEP = 5;

export function renderAudio(src: string, host: HTMLElement, name: string, saved: ReadState, onChange: () => void): Viewer {
  const wrap = el('div', 'center');
  const card = el('div', 'media-card');
  const art = el('div', 'media-art');
  art.append(icon('music'));
  const audio = el('audio');
  audio.controls = true;
  audio.preload = 'metadata';
  audio.src = src;
  card.append(art, el('h1', 'media-name', name), audio, el('p', 'media-hint', `空格 播放 / 暂停 · ← → 后退 / 前进 ${STEP} 秒`));
  wrap.append(card);
  host.append(wrap);

  audio.addEventListener('loadedmetadata', () => {
    const t = saved.time ?? 0;
    if (t > 0 && t < audio.duration - 1) audio.currentTime = t;
  }, { once: true });
  audio.addEventListener('error', () => {
    card.replaceChildren(el('h1', '', '无法播放这个音频'), el('p', '', `VS Code 解码不了「${name}」的编码格式。`));
    card.classList.add('message');
  });
  for (const type of ['pause', 'seeked', 'ended']) audio.addEventListener(type, onChange);
  let lastSave = 0;
  audio.addEventListener('timeupdate', () => {
    if (Math.abs(audio.currentTime - lastSave) < 5) return;
    lastSave = audio.currentTime;
    onChange();
  });

  // 焦点在播放器上时它自己处理按键，不重复响应
  const onKey = (e: KeyboardEvent) => {
    if (e.target === audio || e.metaKey || e.ctrlKey || e.altKey) return;
    if (e.key === ' ') audio.paused ? void audio.play() : audio.pause();
    else if (e.key === 'ArrowLeft') audio.currentTime = Math.max(0, audio.currentTime - STEP);
    else if (e.key === 'ArrowRight') audio.currentTime = Math.min(audio.duration || Infinity, audio.currentTime + STEP);
    else return;
    e.preventDefault();
  };
  window.addEventListener('keydown', onKey);

  return {
    state: () => ({ time: audio.currentTime }),
    destroy() {
      window.removeEventListener('keydown', onKey);
      audio.pause();
      audio.removeAttribute('src');
      audio.load();
    },
  };
}
