export type ScaleMode = 'fitToScreen' | 'fitWidth' | 'custom'

export interface PptxPlayerOptions {
  /** 总页数 */
  slideCount: number
  /** 初始激活页索引（默认 0） */
  initialSlide?: number
  /** 初始缩放模式（默认 fitToScreen） */
  scaleMode?: ScaleMode
  /** 自定义缩放比例（仅在 scaleMode 为 custom 时生效） */
  customScale?: number
}

export interface ScaleChangeEvent {
  mode: ScaleMode
  value: number
}

export type PlayerEvent = 'slideChange' | 'scaleChange' | 'destroy'

type PlayerEventCallback =
  | ((slideIndex: number) => void)
  | ((scale: ScaleChangeEvent) => void)
  | (() => void)

/**
 * PPTX Web Player 播放运行时核心。
 *
 * 职责：
 * - 当前页管理（active slide index）
 * - 导航控制（next / prev / first / last / goToSlide）
 * - 键盘控制（ArrowLeft / ArrowRight / PageUp / PageDown / Home / End / Space）
 * - 视口缩放模式（fitToScreen / fitWidth / custom）
 *
 * 本类不依赖任何 UI 框架，通过事件回调通知状态变更。
 * Vue 集成见 composables/usePlayer.ts。
 */
export class PptxPlayer {
  private _slideCount: number
  private _activeSlideIndex: number
  private _listeners = new Map<PlayerEvent, Set<PlayerEventCallback>>()

  // Keyboard
  private _keyboardTarget: HTMLElement | Document | null = null
  private _keyboardHandler: ((event: Event) => void) | null = null

  // Viewport
  private _scaleMode: ScaleMode
  private _customScale: number

  constructor(options: PptxPlayerOptions) {
    this._slideCount = options.slideCount
    this._activeSlideIndex = this._clampSlide(options.initialSlide ?? 0)
    this._scaleMode = options.scaleMode ?? 'fitToScreen'
    this._customScale = options.customScale ?? 1
  }

  // ── 只读状态 ──

  get activeSlideIndex(): number {
    return this._activeSlideIndex
  }

  get slideCount(): number {
    return this._slideCount
  }

  get canGoNext(): boolean {
    return this._activeSlideIndex < this._slideCount - 1
  }

  get canGoPrevious(): boolean {
    return this._activeSlideIndex > 0
  }

  get scaleMode(): ScaleMode {
    return this._scaleMode
  }

  get scale(): number {
    return this._scaleMode === 'custom' ? this._customScale : 1
  }

  // ── 导航 ──

  /**
   * 更新总页数并在当前页越界时自动修正。
   * 在解析完成或 slide 列表变化时调用。
   */
  updateSlideCount(count: number): void {
    this._slideCount = count
    if (this._activeSlideIndex >= count) {
      this._goToSlideInternal(Math.max(0, count - 1))
    }
  }

  goToSlide(index: number): void {
    this._goToSlideInternal(this._clampSlide(index))
  }

  next(): void {
    this.goToSlide(this._activeSlideIndex + 1)
  }

  previous(): void {
    this.goToSlide(this._activeSlideIndex - 1)
  }

  first(): void {
    this.goToSlide(0)
  }

  last(): void {
    this.goToSlide(this._slideCount - 1)
  }

  // ── 视口缩放 ──

  setScaleMode(mode: ScaleMode): void {
    if (this._scaleMode !== mode) {
      this._scaleMode = mode
      this._emit('scaleChange', { mode, value: this.scale })
    }
  }

  setCustomScale(scale: number): void {
    this._customScale = scale
    if (this._scaleMode === 'custom') {
      this._emit('scaleChange', { mode: 'custom', value: scale })
    }
  }

  // ── 键盘 ──

  /**
   * 在目标元素上绑定键盘事件。默认绑定到 document。
   * 重复调用会先解绑上次绑定。
   */
  attachKeyboard(target: HTMLElement | Document = document): void {
    this.detachKeyboard()
    this._keyboardTarget = target

    this._keyboardHandler = (event: Event) => {
      const keyEvent = event as KeyboardEvent
      if (PptxPlayer.isInteractiveTarget(keyEvent.target)) return

      switch (keyEvent.key) {
        case 'ArrowLeft':
        case 'PageUp':
          event.preventDefault()
          this.previous()
          break
        case 'ArrowRight':
        case 'PageDown':
        case ' ':
          event.preventDefault()
          this.next()
          break
        case 'Home':
          event.preventDefault()
          this.first()
          break
        case 'End':
          event.preventDefault()
          this.last()
          break
      }
    }

    this._keyboardTarget.addEventListener('keydown', this._keyboardHandler)
  }

  detachKeyboard(): void {
    if (this._keyboardHandler && this._keyboardTarget) {
      this._keyboardTarget.removeEventListener('keydown', this._keyboardHandler)
    }
    this._keyboardHandler = null
    this._keyboardTarget = null
  }

  // ── 事件 ──

  on(event: PlayerEvent, callback: PlayerEventCallback): void {
    if (!this._listeners.has(event)) {
      this._listeners.set(event, new Set())
    }
    this._listeners.get(event)!.add(callback)
  }

  off(event: PlayerEvent, callback: PlayerEventCallback): void {
    this._listeners.get(event)?.delete(callback)
  }

  // ── 生命周期 ──

  destroy(): void {
    this.detachKeyboard()
    this._emit('destroy')
    this._listeners.clear()
  }

  // ── Private ──

  private _goToSlideInternal(index: number): void {
    if (index !== this._activeSlideIndex) {
      this._activeSlideIndex = index
      this._emit('slideChange', index)
    }
  }

  private _clampSlide(index: number): number {
    if (this._slideCount <= 0) return 0
    return Math.max(0, Math.min(index, this._slideCount - 1))
  }

  private _emit(event: PlayerEvent, ...args: unknown[]): void {
    this._listeners.get(event)?.forEach((callback) => {
      try {
        ;(callback as (...a: unknown[]) => void)(...args)
      } catch {
        // 吞掉单个监听器的异常，避免影响其他监听器
      }
    })
  }

  /**
   * 判断事件目标是否为交互式元素（输入框、按钮等）。
   * 在这些元素上触发的键盘事件不应被播放器拦截。
   */
  static isInteractiveTarget(target: EventTarget | null): boolean {
    return (
      target instanceof HTMLElement &&
      ['BUTTON', 'INPUT', 'TEXTAREA', 'SELECT', 'A'].includes(target.tagName)
    )
  }
}
