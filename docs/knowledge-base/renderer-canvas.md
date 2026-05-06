# Canvas Renderer 记录

## Canvas renderer 只消费 normalized model

### 现象

Canvas 预览出现位置、层级、背景或图片裁剪问题时，容易直接在 Canvas 绘制分支里补特殊逻辑。

### 原因

当前架构要求 SVG 与 Canvas renderer 都只消费 `@pptx-player/core` 输出的 normalized model，不读取 raw OOXML。Canvas 后端拿到的是 `Presentation`、`Slide`、`SlideElement` 和已加载的 `mediaBitmaps`。

### 解决方案

排查 Canvas 问题时按同一条链路定位：

1. core model 是否已有 `transform`、`zIndex`、`fill`、`line`、`crop`、`opacity`、`background`
2. layout 是否已产出正确 text lines
3. Canvas renderer 是否把 normalized 字段映射到 `CanvasRenderingContext2D`

不要让 Canvas renderer 解析 XML 或重做 master/layout 继承。

### 关联文件

- `packages/renderer-canvas/src/index.ts`
- `packages/renderer-canvas/src/index.spec.ts`
- `packages/core/src/model/Presentation.ts`
- `packages/layout/src/index.ts`

## Canvas 背景、层级和透明度约定

### 现象

同一页在 SVG 与 Canvas 之间可能出现背景遮挡、元素顺序或透明度不一致。

### 原因

Canvas 是立即模式绘制，绘制顺序就是最终层级；透明度依赖 `globalAlpha` 的保存和恢复。如果背景或元素顺序没有与 SVG 对齐，会直接产生视觉差异。

### 解决方案

当前约定：

- 背景永远先绘制。
- 图片背景绘制到整页，并复用图片 crop 逻辑。
- explicit `fill: { type: 'none' }` 背景保持透明，不填默认色。
- 普通元素按 `zIndex` 升序绘制，同 `zIndex` 保持原数组稳定顺序。
- hidden 或 `opacity <= 0` 的元素跳过。
- 半透明元素通过 `context.globalAlpha *= element.opacity` 实现，并用 `save()` / `restore()` 隔离。

### 关联文件

- `packages/renderer-canvas/src/index.ts`
- `packages/renderer-canvas/src/index.spec.ts`
- `packages/renderer-svg/src/index.ts`

## Canvas 图片裁剪依赖 bitmap 原始尺寸

### 现象

带 crop 的图片在 Canvas 中如果只使用 5 参数 `drawImage()`，会把整张图拉伸到目标框，和 PPT 可见区域不一致。

### 原因

PPT 的图片 crop 来自 normalized `ImageCrop`，表示源图四边裁掉的比例。Canvas 需要源 bitmap 的宽高才能换算出 9 参数 `drawImage()` 的 source rectangle。

### 解决方案

当前 Canvas renderer：

- 无 crop 时使用 5 参数 `drawImage(bitmap, x, y, width, height)`。
- 有 crop 且 bitmap 尺寸可用时，计算 `sx/sy/sw/sh` 后使用 9 参数 `drawImage()`。
- bitmap 尺寸不可用时安全降级为整图绘制。
- 缺图时绘制蓝色占位框和 imagePart/relationshipId，便于调试 media 加载问题。

### 关联文件

- `packages/renderer-canvas/src/index.ts`
- `packages/core/src/ooxml/presentation/parseImage.ts`
- `packages/core/src/ooxml/presentation/parseImageFill.ts`

## SVG/Canvas 差异排查顺序

### 现象

视觉回归中同一页可能只有 SVG 失败、只有 Canvas 失败，或两个后端同时失败。

### 原因

两个 renderer 共用 normalized model 和 layout，但输出后端能力不同。Canvas 目前没有完整实现 SVG 已有的部分声明式能力，例如 connector marker 和 underline。

### 解决方案

按下面顺序判断归属：

1. 两个后端同时缺同一元素或同一字段：优先查 core/parser。
2. 两个后端文字行位置都错：优先查 `packages/layout`。
3. SVG 正确、Canvas 错：查 Canvas 绘制 API 映射、bitmap 尺寸、`globalAlpha` 和绘制顺序。
4. Canvas 正确、SVG 错：查 SVG 字符串输出、viewBox、attribute escaping 和 marker/defs。
5. 与 WPS/Office 对比时，用 `pnpm visual:compare` 或 `pnpm visual:compare:batch` 产出 actual/diff/report 后再定位。

### 关联文件

- `packages/renderer-canvas/src/index.ts`
- `packages/renderer-svg/src/index.ts`
- `packages/layout/src/index.ts`
- `scripts/visual-compare.mjs`
