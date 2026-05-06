# Renderer / Core Model 对接记录

## 先看 core 模型是否已经表达问题，再看 renderer

### 现象

页面上元素渲染不对时，很容易第一反应就改 renderer，但不少问题其实在 core 解析层已经发生了信息丢失。

### 原因

renderer 只能消费 core 暴露出来的模型。如果 `transform`、`fill`、`textBody`、`bullet`、`diagnostics` 没进入模型，renderer 再精细也无从还原。

### 解决方案

排查顺序始终是：

1. 先看 `parsePptx()` 输出模型
2. 再看 renderer 是否消费该字段
3. 最后才改样式或布局实现

### 关联文件

- `packages/core/src/parser/parsePptx.ts`
- `packages/core/src/model/Presentation.ts`
- `packages/core/src/__tests__/parsePptx.spec.ts`

## SlideElementBase 是 renderer 的最低输入面

### 现象

无论是 text、shape、image 还是 connector，renderer 排布时都离不开一组共通字段。

### 原因

这些字段决定了渲染对象的空间位置、层级和可见性，是所有元素的共同语义底座。

### 解决方案

renderer 侧优先围绕 `SlideElementBase` 建立消费约定，尤其关注：

- `transform`
- `visible`
- `opacity`
- `zIndex`
- `source`
- `diagnostics`

缺这些字段时，优先回 core 解析层补齐。

### 关联文件

- `packages/core/src/model/Presentation.ts`
- `packages/core/src/__tests__/parsePptx.spec.ts`

## 扁平 text 只适合粗消费，精细渲染要读 textBody

### 现象

如果 renderer 只拿 `element.text` 去画文本，能显示内容，但列表、run 样式、换行、字段文本细节会很快失真。

### 原因

`text` 是为搜索、预览、简单 UI 提供的扁平输出；真正的样式和结构都在 `textBody`。

### 解决方案

- 简单展示可使用 `text`
- 需要接近 PPT 视觉还原时，必须消费 `textBody.paragraphs[].runs[]` 及 paragraph/run style
- 列表符号要读 `paragraph.style?.bullet`

### 关联文件

- `packages/core/src/model/Presentation.ts`
- `packages/core/src/ooxml/presentation/parseShape.ts`
- `packages/core/src/ooxml/presentation/parseTextBody.ts`

## diagnostics 不是噪音，要带着它调试 renderer

### 现象

某些元素“看起来像渲染丢了”，但实际上 core 已经告诉你关系缺失、资源找不到或解析降级了。

### 原因

当前项目会把包加载、关系解析、元素解析中的异常沉淀为 `diagnostics`，并尽量附带 part、slide、element 上下文。

### 解决方案

renderer 调试时同时查看：

- `result.diagnostics`
- `slide.diagnostics`
- `element.diagnostics`

先确认是“数据没来”，还是“renderer 没画”。

### 关联文件

- `packages/core/src/diagnostics/context.ts`
- `packages/core/src/package/PptxPackage.ts`
- `packages/core/src/ooxml/presentation/parseImage.ts`
- `packages/core/src/__tests__/parsePptx.spec.ts`

## zIndex 与元素遍历顺序要一起验证

### 现象

元素遮挡关系错误时，问题可能不是 CSS/SVG 叠放逻辑，而是解析出来的顺序和层级预期不一致。

### 原因

slide 元素的视觉叠放通常依赖解析顺序和显式 `zIndex`。只看其中一个维度会误判。

### 解决方案

检查：

- core 输出数组顺序
- 每个元素的 `index`
- 每个元素的 `zIndex`
- renderer 是否按 `zIndex` 升序绘制，并在同 `zIndex` 时保持原数组稳定顺序
- background 是否始终先于普通元素绘制

### 关联文件

- `packages/core/src/model/Presentation.ts`
- `packages/core/src/__tests__/parsePptx.spec.ts`

## Week 3 MVP 先建立 renderer 边界，再谈更高保真

### 现象

在项目早期，`packages/app` 里直接写 SVG 预览逻辑推进很快，但 renderer 包会长期空转，最终导致 parser、renderer、app 的职责边界混乱。

### 原因

如果 app 同时负责：

- 文件上传
- PPTX 解析
- slide 切换
- SVG 元素级绘制

那么 renderer 包即使存在，也不会成为真实依赖点，后续很难演进出独立的渲染层。

### 解决方案

Week 3 MVP 先把正式边界立起来：

- `@pptx-player/renderer-svg` 接收 normalized model
- `@pptx-player/renderer-canvas` 也直接消费同一份 normalized model
- renderer 负责各自后端的单位映射和绘制细节
- `packages/app` 只负责解析、翻页、media URL 管理和渲染结果挂载

当前 renderer 的入口约定为 `renderSlideToSvg()`，输入：

- `presentation.width`
- `presentation.height`
- `slide`
- `mediaUrls`

先支持这些静态元素：

- slide background
- text
- image
- basic shape
- connector

其中 `renderer-canvas` 当前也已具备同等级别的静态骨架，并与 SVG 共享 `packages/layout` 的文本布局入口：

- 直接消费 `textBody.paragraphs[].runs[]`
- 复用 `layoutTextElement()` 处理换行、run、bullet、auto-number、缩进和段落间距
- 依赖 canvas `measureText()` 做 run 宽度测量
- 与 SVG 的差异主要来自具体绘制后端能力，而不是 normalized model 不一致

### 关联文件

- `packages/renderer-svg/src/index.ts`
- `packages/app/src/App.vue`
- `packages/app/package.json`
- `packages/app/tsconfig.app.json`

## app 不应继续掌握元素级 SVG 细节

### 现象

如果 app 中保留大量 `v-if="element.type === ..."` 的 SVG 分支，随着支持元素增多，`App.vue` 会快速膨胀，且每新增一种渲染能力都要跨包修改。

### 原因

这类逻辑本质属于 renderer，而不是示例应用层。

### 解决方案

把以下内容收口到 `renderer-svg`：

- marker defs
- shape / ellipse / connector 的 SVG 拼装
- line dash / marker 映射
- image fallback 占位绘制
- preview SVG 的字符串拼接

app 中保留：

- `parsePptx()` 调用
- active slide 状态
- `mediaUrls` 生命周期
- `v-html` 挂载 renderer 输出

### 关联文件

- `packages/renderer-svg/src/index.ts`
- `packages/app/src/App.vue`

## 图片 crop / alpha 要在 core 归一化后由 renderer 消费

### 现象

PPT 中裁剪过或带透明度的图片，如果 renderer 只按完整 bitmap 铺到目标框，会出现可见区域过大、透明度不对等问题。

### 原因

图片裁剪语义来自 `p:blipFill/a:srcRect`，透明度常来自 `a:blip/a:alphaModFix` 或 `a:alpha`。renderer 不应该读取 raw XML；如果 core 没把这些字段归一化，SVG 和 Canvas 后端都无法可靠复用。

### 解决方案

当前 core 将图片语义归一化到 model：

- `ImageElement.crop` 使用 `0..1` 的 `left/top/right/bottom`
- 图片 alpha 合并到元素级 `opacity`
- 非法 crop 或 alpha 输出 `IMAGE_CROP_INVALID` / `IMAGE_ALPHA_INVALID` diagnostics，并安全降级

renderer 的消费方式：

- SVG 无 crop 时继续输出普通 `<image>`；有 crop 时输出嵌套 `<svg viewBox>`，用 normalized viewBox 表达裁剪源区域
- Canvas 无 crop 时用 5 参数 `drawImage()`；有 crop 且 bitmap 尺寸可用时用 9 参数 `drawImage()` 绘制源矩形到目标矩形
- external image 仍只做降级，不在本轮下载或跨域加载

### 关联文件

- `packages/core/src/model/Presentation.ts`
- `packages/core/src/ooxml/presentation/parseImage.ts`
- `packages/core/src/diagnostics/codes.ts`
- `packages/renderer-svg/src/index.ts`
- `packages/renderer-canvas/src/index.ts`

## 背景图也必须从 core normalized model 进入 renderer

### 现象

WPS/Office 中看起来像整页底图的内容，可能不是普通 `p:pic`，而是 `p:bg/p:bgPr/a:blipFill` 背景图。如果只收集 slide elements，会出现 diagnostics 正常但页面大块背景缺失。

### 原因

slide background 与 shape/image element 是不同 OOXML 路径。renderer 不应该读取 raw XML，因此背景图必须先由 core 解析成 `SlideBackground`，并把本地 `imagePart` 纳入 `parsePptx().media`。

### 解决方案

当前 core 将背景归一化为：

- solid/noFill 背景：`{ type: 'fill', fill }`
- 图片背景：`{ type: 'image', fill: ImageFill }`
- `ImageFill` 复用普通图片的 relationship、crop、opacity 解析
- external 背景图只保留 relationship 并降级，不下载

renderer 的消费方式：

- SVG 在元素前先输出 background；图片背景按整页 transform 绘制，并复用 crop viewBox 逻辑
- Canvas 在元素前先绘制 background；图片背景绘制到整页，并复用 cropped `drawImage()` 逻辑
- explicit noFill 背景保持透明，不强制填充默认背景
- 未声明背景时才回退到默认背景色，SVG 默认矩形不添加非 PPT 圆角
- 背景图缺 media 时回退到默认背景色

### 关联文件

- `packages/core/src/model/Presentation.ts`
- `packages/core/src/ooxml/presentation/parseSlideBackground.ts`
- `packages/core/src/ooxml/presentation/parseImageFill.ts`
- `packages/core/src/parser/parsePptx.ts`
- `packages/renderer-svg/src/index.ts`
- `packages/renderer-canvas/src/index.ts`

## SVG 与 Canvas renderer 的当前能力差异

### 现象

同一份 slide normalized model 交给 SVG 与 Canvas 后端时，基础结构一致，但少量视觉细节不会完全一致。

### 原因

两个 renderer 都只消费 core normalized model，不读取原始 XML；差异来自输出目标不同：SVG 输出可声明的 DOM / marker / text-decoration，Canvas 则需要手动画路径和文本装饰。

### 当前矩阵

| 能力 | renderer-svg | renderer-canvas |
| --- | --- | --- |
| 输入模型 | `Presentation` + `Slide` + `mediaUrls` | `Presentation` + `Slide` + `mediaBitmaps` |
| text layout | 复用 `layoutTextElement()`，用估算宽度测量 | 复用 `layoutTextElement()`，用 `context.measureText()` 测量 |
| run 样式 | color / fontFace / fontSize / bold / italic / underline | color / fontFace / fontSize / bold / italic |
| visibility / opacity | 跳过 hidden / 透明元素，半透明元素用 `<g opacity>` | 跳过 hidden / 透明元素，半透明元素乘到 `globalAlpha` |
| shape fill / line | solid/noFill、fill opacity、stroke opacity、dash | solid/noFill、fill opacity、stroke opacity、dash |
| connector marker | 支持 headEnd / tailEnd 的基础 marker | 暂不绘制 arrowhead marker |
| 输出形态 | SVG 字符串，可由 app 用 `v-html` 挂载 | 直接绘制到 `CanvasRenderingContext2D` |

### 解决方案

调试时按能力矩阵判断问题归属：

- SVG 有箭头而 Canvas 没箭头，当前是 Canvas renderer 能力差异，不是 parser 丢了 `LineStyle.headEnd/tailEnd`。
- SVG 有 underline 而 Canvas 没 underline，当前是 Canvas 文本装饰未实现，不是 `TextStyle.underline` 没解析。
- 两个 renderer 都缺同一字段时，优先回 core model 和 parser 查。
- 只有某个 renderer 缺视觉表现时，再改对应 renderer。

### 关联文件

- `packages/renderer-svg/src/index.ts`
- `packages/renderer-canvas/src/index.ts`
- `packages/layout/src/index.ts`
- `packages/core/src/model/Presentation.ts`

## Workspace 包接入时，先处理 TypeScript 引用链

### 现象

新建 workspace 包后，运行 type-check 时 app 可能报找不到模块，或者跨包类型检查把上游已知兼容问题重新暴露出来。

### 原因

当前仓库处于源码直连阶段，不是所有包都已走完整构建产物链路；因此：

- app 侧需要显式 TS path 映射到源码入口
- 新包跨依赖 core 时，可能继承 core 的类型兼容补丁需求

### 解决方案

接入新的 workspace renderer 包时优先检查：

- `package.json` dependency 是否已声明
- app 的 `tsconfig.app.json` 是否加了 path 映射
- 新包是否需要补齐与 core 一致的类型 shim（例如 `txml`）

### 关联文件

- `packages/app/tsconfig.app.json`
- `packages/renderer-svg/src/txml.d.ts`
- `packages/core/src/txml.d.ts`

## renderer-svg 文本渲染要先做单位归一化，再谈排版

### 现象

文本已经能显示，但常见问题会集中爆发：

- 字号异常大
- 行距过密或过松
- 段前段后距离明显不对
- 列表缩进和续行位置错乱

### 原因

`packages/core` 当前保留的是更接近 OOXML 原始语义的文本数值，而不是浏览器可直接使用的 CSS/SVG 单位。例如：

- `TextStyle.fontSize` 是 `1/100 pt`
- `ParagraphStyle.marginLeft` / `marginRight` / `indent` 是 EMU
- `TextSpacing.points` 是 `1/100 pt`
- `TextSpacing.percent` 是 `100000 = 100%`

如果 renderer 直接把这些值写进 SVG：

- `1800` 会被误当成 `1800px`，而不是 `18pt`
- `457200` 会被误当成像素，而不是约 `36px`
- `150000` 会被误当成极大行高，而不是 `1.5x`

### 解决方案

在 `renderer-svg` 中先建立统一换算，再进入排版逻辑：

- 字号：`fontSize / 100`
- EMU：`value / 12700`
- 百分比行距：`percent / 100000`
- point spacing：先还原 point 语义，再映射到当前 SVG 字号逻辑

建议所有文本布局辅助函数都只消费“已归一化后的 SVG 数值”，不要让原始 OOXML 单位继续泄漏到后续渲染步骤。

### 关联文件

- `packages/renderer-svg/src/index.ts`
- `packages/core/src/model/Presentation.ts`
- `packages/core/src/ooxml/presentation/parseTextBody.ts`
- `packages/renderer-svg/src/index.spec.ts`

## Week 3 renderer-svg 文本支持边界

### 现象

Week 3 之前，`renderer-svg` 对 text 元素基本只是占位展示：直接取 `element.text` 前若干字符输出为单行 `<text>`。

### 原因

这种方式能快速证明“文本元素存在”，但无法表达 normalized text model 里的真实结构和样式：

- paragraph 分段
- run 粒度样式
- bullet
- 多行
- 段落间距与缩进

### 解决方案

Week 3 当前实现已经升级为真正消费 `textBody`：

- 按 `textBody.paragraphs[]` 渲染多段文本
- 按 `paragraph.runs[]` 输出 run 级 `<tspan>`
- 支持 `bold` / `italic` / `underline` / `fontFace` / `color` / `fontSize`
- 支持 paragraph 级 `bullet`
- 支持 `a:br` 带来的多行拆分
- 支持基础 `align` 映射：左 / 中 / 右
- 支持基础 `marginLeft` / `marginRight` / `indent`
- 支持基础 `lineSpacing` / `spaceBefore` / `spaceAfter`

当前仍然属于“结构正确 + 基础视觉接近”的阶段，而不是最终高保真布局。

### 已知限制

当前 `renderer-svg` 仍未覆盖或仅做了近似处理：

- bullet 宽度没有做真实字形测量，只按文本前缀处理
- `buSzPct` / `buSzTx` 还没有完整语义化还原
- `rtl` 仅保留在模型中，renderer 尚未做专门布局
- paragraph 对齐当前按整段基础映射，未做更细粒度排版校正
- tab stop、复杂换行、fallback font、CJK line breaking 仍未进入 layout 层
- 当前没有独立 text layout engine，行宽裁剪、自动换行、真实字体测量仍是后续工作

### 关联文件

- `packages/renderer-svg/src/index.ts`
- `packages/renderer-svg/src/index.spec.ts`
- `packages/app/src/App.vue`
- `packages/core/src/model/Presentation.ts`

## renderer 不负责 master/layout 样式继承

### 现象

同一个 placeholder 在 SVG 与 Canvas 中都缺少 fill、line 或默认文字样式时，容易误以为两个 renderer 都漏了样式处理。

### 原因

当前架构要求 renderer 只消费 normalized model，不读取 raw XML，也不重新执行 theme/master/layout/slide 的继承链。Week 2 的最小样式继承已经在 core parser 中完成：slide element 输出时应尽量带上合并后的 fill、line、textBody run style 与 background。

### 解决方案

排查样式继承问题时先看 core 输出：

- `presentation.slideMasters[].defaults`
- `presentation.slideLayouts[].defaults`
- `slide.background`
- `element.fill` / `element.line`
- `element.textBody.paragraphs[].runs[].style`
- `STYLE_INHERITANCE_INCOMPLETE` diagnostics

只有 core 模型中字段已存在但某个后端没画出来时，才修改对应 renderer。

### 关联文件

- `packages/core/src/ooxml/presentation/parseSlideMaster.ts`
- `packages/core/src/ooxml/presentation/parseSlide.ts`
- `packages/core/src/ooxml/presentation/parseShape.ts`
- `packages/renderer-svg/src/index.ts`
- `packages/renderer-canvas/src/index.ts`

## 遇到渲染异常时的排查清单

### 现象

常见表现有：元素不显示、位置错乱、样式缺失、列表不对、图片不出来、明明有数据却渲染为空。

### 原因

问题通常跨 core 与 renderer，两边任意一层丢语义都会表现成“渲染问题”。

### 解决方案

建议按下面顺序排查：

1. 用 `parsePptx()` 先看 presentation / slide / element 模型是否完整
2. 检查目标元素的 `type`、`transform`、`visible`、`opacity`、`zIndex`
3. 文本问题看 `textBody`，不要只看扁平 `text`
4. 图片问题看 relationship 解析结果和 `media` 是否存在
5. 列表问题看 `paragraph.style.bullet`
6. 如果存在 diagnostics，优先理解 diagnostics 再动 renderer
7. 只有当 core 模型正确时，再修改 renderer 实现
8. 对照 WPS/Office 视觉结果时，使用 `pnpm visual:compare` 生成 actual / diff / report，不再只依赖人工肉眼截图判断

### 关联文件

- `packages/core/src/parser/parsePptx.ts`
- `packages/core/src/model/Presentation.ts`
- `packages/core/src/__tests__/parsePptx.spec.ts`
