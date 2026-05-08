import { computed, onBeforeUnmount, onMounted, ref, type Ref } from 'vue'
import { PptxPlayer, type PptxPlayerOptions, type ScaleMode } from '@pptx-player/player'

export interface UsePlayerReturn {
  activeSlideIndex: Ref<number>
  slideCount: Ref<number>
  canGoNext: Ref<boolean>
  canGoPrevious: Ref<boolean>
  scaleMode: Ref<ScaleMode>
  scale: Ref<number>
  goToSlide: (index: number) => void
  next: () => void
  previous: () => void
  first: () => void
  last: () => void
  updateSlideCount: (count: number) => void
  setScaleMode: (mode: ScaleMode) => void
  setCustomScale: (scale: number) => void
}

/**
 * 将 PptxPlayer 包装为 Vue 3 响应式 composable。
 *
 * 用法：
 * ```ts
 * const player = usePlayer({ slideCount: 10 })
 * ```
 *
 * 组件挂载时自动绑定键盘事件，卸载时自动调用 player.destroy() 清理。
 */
export function usePlayer(options: PptxPlayerOptions): UsePlayerReturn {
  const player = new PptxPlayer(options)

  const activeSlideIndex = ref(player.activeSlideIndex)
  const scaleMode = ref<ScaleMode>(player.scaleMode)
  const scaleValue = ref(player.scale)

  player.on('slideChange', (index: number) => {
    activeSlideIndex.value = index
  })

  player.on('scaleChange', ({ mode, value }: { mode: ScaleMode; value: number }) => {
    scaleMode.value = mode
    scaleValue.value = value
  })

  onMounted(() => {
    player.attachKeyboard()
  })

  onBeforeUnmount(() => {
    player.destroy()
  })

  return {
    activeSlideIndex,
    slideCount: computed(() => player.slideCount),
    canGoNext: computed(() => player.canGoNext),
    canGoPrevious: computed(() => player.canGoPrevious),
    scaleMode,
    scale: computed(() => player.scale),
    goToSlide: (index: number) => player.goToSlide(index),
    next: () => player.next(),
    previous: () => player.previous(),
    first: () => player.first(),
    last: () => player.last(),
    updateSlideCount: (count: number) => player.updateSlideCount(count),
    setScaleMode: (mode: ScaleMode) => player.setScaleMode(mode),
    setCustomScale: (scale: number) => player.setCustomScale(scale),
  }
}
