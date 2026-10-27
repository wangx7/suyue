import type { Shell } from '../shell';
import type { ReadState } from '../store';
import type { Zoomable } from '../zoom';

/** 文档里的相对路径（图片、指向其他文件的链接），按文档网址解析 */
export interface Linker {
  /** 返回可直接用作 src 的地址，找不到时返回 null */
  asset(path: string): Promise<string | null>;
  /** 打开链接指向的文件；打不开时返回 false */
  open(path: string): boolean;
}

export interface ViewerContext {
  /** 已挂到页面上的渲染容器（绝对定位铺满正文区） */
  host: HTMLElement;
  shell: Shell;
  name: string;
  /** 上次的阅读位置，没有则为空对象 */
  saved: ReadState;
  /** 从网址打开时才有；本地文件无法读取同目录的其他文件 */
  linker?: Linker;
}

export interface Viewer {
  destroy(): void;
  onTheme?(dark: boolean): void;
  zoom?: Zoomable;
  /** 当前阅读状态，用于下次打开时恢复 */
  state?(): ReadState;
  /** 主滚动容器：用于“回到顶部”和滚动停止后保存位置 */
  scroller?: HTMLElement;
  /** 自带搜索（Canvas 表格浏览器搜不到），有则接管 Ctrl+F */
  openSearch?(): void;
}

export type Renderer = (data: ArrayBuffer, ctx: ViewerContext) => Promise<Viewer>;
