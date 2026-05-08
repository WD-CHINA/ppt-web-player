# Player Runtime

## 概述

`@pptx-player/player` 是 PPTX Web Player 的播放运行时核心包。

## 架构设计

```
packages/player/src/
  PptxPlayer.ts    # 播放运行时核心类（框架无关）
  index.ts          # 公开导出
```

Vue 集成 composable 位于 `packages/app/src/composables/usePlayer.ts`。

## 职责边界

- **PptxPlayer 负责**：active slide 管理、导航控制、键盘事件、视口缩放模式
- **PptxPlayer 不负责**：文件解析、渲染、UI 布局、事件总线之外的任何框架耦合

## PptxPlayer API

```ts
import { PptxPlayer } from '@pptx-player/player'

const player = new PptxPlayer({
  slideCount: 10,           // 总页数
  initialSlide: 0,          // 初始页（可选，默认 0）
  scaleMode: 'fitToScreen', // 缩放模式（可选）
})

// 导航
player.goToSlide(3)
player.next()
player.previous()
player.first()
player.last()

// 更新页数（解析完成后调用）
player.updateSlideCount(20)

// 键盘（默认绑定 document）
player.attachKeyboard()
player.detachKeyboard()

// 事件
player.on('slideChange', (index) => { ... })
player.on('scaleChange', ({ mode, value }) => { ... })

// 销毁
player.destroy()
```

## Vue 集成

```ts
import { usePlayer } from '@/composables/usePlayer'

const {
  activeSlideIndex,  // Ref<number>
  slideCount,        // Ref<number>
  canGoNext,         // Ref<boolean>
  canGoPrevious,     // Ref<boolean>
  goToSlide,         // (index: number) => void
  next, previous, first, last,
  updateSlideCount,
} = usePlayer({ slideCount: 0 })
```

组件挂载时自动绑定键盘事件，卸载时自动解绑。

## 缩放模式

| 模式 | 说明 |
|------|------|
| `fitToScreen` | 自适应容器（CSS 实现，scale 返回 1） |
| `fitWidth` | 适应宽度 |
| `custom` | 自定义缩放比例 |

## 当前限制

- 不支持逐 step 动画播放（Phase 2 规划）
- 不支持触控事件
- 缩放模式的实际 CSS/Canvas 应用逻辑仍在 App.vue 中
