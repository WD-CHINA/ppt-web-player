# Visual Regression / WPS 对比记录

## WPS 参考图要作为渲染回归的外部基线

### 现象

人工对照 WPS 截图时，只能主观判断“像不像”，很难稳定追踪 SVG 与 Canvas 哪个后端退化了。

### 原因

当前项目同时存在 parser、layout、renderer-svg、renderer-canvas 多层输出。只看 DOM 或 diagnostics 能证明模型存在，但不能证明最终视觉接近 WPS。

### 解决方案

新增根命令：

```sh
pnpm visual:compare -- --reference <wps.png> --sample "Fixture 4b00" --slide 8 --renderer svg
pnpm visual:compare -- --reference <wps.png> --sample "Fixture 4b00" --slide 8 --renderer canvas
```

脚本行为：

1. 连接正在运行的 app，例如 `http://localhost:5175/`
2. 点击指定 sample
3. 翻到指定 slide
4. 截取 SVG 或 Canvas 预览区域
5. 将 WPS 参考图缩放到当前预览截图尺寸
6. 使用 pixel diff 输出 actual、diff、report

输出默认落在：

```text
reports/visual-regression/
```

报告字段包含：

- `mismatchedPixels`
- `totalPixels`
- `mismatchRatio`
- `referenceWidth` / `referenceHeight`
- `actualWidth` / `actualHeight`
- `actual`
- `diff`

### 使用约定

- 先运行 app：`pnpm dev`
- WPS 参考图由人工从 WPS/Office 导出或截图提供。
- 参考图不要求与浏览器截图同尺寸，脚本会缩放参考图后对比。
- `--pixelThreshold` 控制单像素颜色容忍度，默认 `0.1`。
- `--maxMismatchRatio` 控制整张图失败阈值，默认 `0.05`。
- 当前阶段阈值应偏宽，用来定位大块缺图、层级错误、背景错误，不用于追求像素级 Office 等价。
- 如果 WPS 参考图与实际渲染存在整页底图差异，优先检查 core 是否把 `p:bg/p:bgPr/a:blipFill` 解析为 `slide.background.type === 'image'`，以及 `parsePptx().media` 是否包含该背景图 part。

### 关联文件

- `scripts/visual-compare.mjs`
- `package.json`
- `packages/app/src/App.vue`
- `packages/renderer-svg/src/index.ts`
- `packages/renderer-canvas/src/index.ts`
