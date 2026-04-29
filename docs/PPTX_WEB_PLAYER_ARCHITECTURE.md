# 企业可控 PPTX Web Player 架构设计与计划

## 1. 产品目标

核心目标是在浏览器中直接解析并播放 PPTX，支持企业常见 PPT 的高可用预览、基础播放、权限控制、诊断降级和持续兼容演进。

目标不是完全复刻 PowerPoint / WPS，而是构建一个企业级 PPTX Web 播放内核。

优先保证：

1. 文件能打开。
2. 页面能稳定显示。
3. 常见文字、图片、形状、表格尽量准确。
4. 基础动画可播放。
5. 高级能力可降级。
6. 所有降级可诊断。
7. 支持私有化和安全控制。
8. 支持大文件、持续回归。

## 2. 总体架构

```text
┌─────────────────────────────────────────────┐
│                PPTX Web Player              │
├─────────────────────────────────────────────┤
│                 Player UI                   │
│ 翻页 / 全屏 / 缩略图 / 搜索 / 备注 / 批注 / 水印 │
├─────────────────────────────────────────────┤
│              Player Runtime                 │
│ 播放状态 / 键盘控制 / 预加载 / 动画调度 / 视口缩放 │
├─────────────────────────────────────────────┤
│              Renderer Layer                 │
│ SVG Renderer / Canvas Renderer / DOM Overlay │
├─────────────────────────────────────────────┤
│              Layout Engine                  │
│ 文本布局 / 表格布局 / Group Transform / 图表布局 │
├─────────────────────────────────────────────┤
│              Normalized Model               │
│ Presentation / Slide / Element / Style / Timeline │
├─────────────────────────────────────────────┤
│              OOXML Parser                   │
│ PresentationML / DrawingML / ChartML / Animation │
├─────────────────────────────────────────────┤
│              Package Reader                 │
│ ZIP / ContentTypes / Relationships / Media / XML Cache │
└─────────────────────────────────────────────┘
```

## 3. 模块分层

### 3.1 Package Reader

负责读取 `.pptx` 包结构。

职责：

- 解压 PPTX。
- 读取 `[Content_Types].xml`。
- 读取 `_rels/.rels`。
- 解析 slide、layout、master、theme、media 关系。
- 统一处理相对路径。
- 缓存 XML。
- 缓存媒体文件。
- 记录缺失资源和异常资源。

目录建议：

```text
src/package/
  PptxPackage.ts
  ContentTypes.ts
  Relationship.ts
  RelationshipResolver.ts
  PartName.ts
  MediaStore.ts
```

核心接口：

```ts
class PptxPackage {
  getPart(path: string): Promise<ArrayBuffer | null>
  getXml(path: string): Promise<XmlNode | null>
  getRelationships(partPath: string): Promise<Relationship[]>
  resolveRelationship(partPath: string, rId: string): ResolvedRelationship | null
}
```

设计重点：这一层必须替代所有散落的字符串路径处理，不要到处写 `target.replace('../', 'ppt/')`，必须统一走 `relationshipResolver.resolve(basePart, target)`。

### 3.2 XML Layer

负责将 XML 转成统一节点结构，并提供安全访问 API。

职责：

- XML parse。
- namespace 保留。
- 属性读取。
- 子节点读取。
- 单节点/数组归一化。
- XPath-like 查询。
- 保留节点顺序。
- 提供 source location，便于 diagnostics。

目录建议：

```text
src/xml/
  XmlParser.ts
  XmlNode.ts
  XmlQuery.ts
  namespaces.ts
```

示例 API：

```ts
xml.child(node, 'p:cSld')
xml.children(node, 'a:p')
xml.attr(node, 'r:id')
xml.text(node)
xml.path(node, ['p:sld', 'p:cSld', 'p:spTree'])
```

### 3.3 OOXML Parser

把 OOXML 解析为中间模型，不直接负责渲染。

目录建议：

```text
src/ooxml/
  presentation/
    parsePresentation.ts
    parseSlide.ts
    parseLayout.ts
    parseMaster.ts
    parseNotes.ts
    parseTransition.ts
    parseTiming.ts

  drawing/
    parseTransform.ts
    parseFill.ts
    parseLine.ts
    parseTextBody.ts
    parseParagraph.ts
    parseRun.ts
    parseShape.ts
    parseGeometry.ts
    parseEffects.ts
    parseImage.ts
    parseGroup.ts

  table/
    parseTable.ts
    parseTableStyle.ts
    parseCell.ts

  chart/
    parseChart.ts
    parseSeries.ts
    parseAxis.ts
    parseLegend.ts

  media/
    parseVideo.ts
    parseAudio.ts

  math/
    parseOMath.ts

  diagram/
    parseSmartArt.ts
```

## 4. 标准化模型设计

PPTX 原始 XML 非常复杂，播放器不应该直接消费 XML，而应该消费标准化模型。

### 4.1 Presentation

```ts
interface Presentation {
  id: string
  width: number
  height: number
  theme: Theme
  slides: Slide[]
  masters: SlideMaster[]
  layouts: SlideLayout[]
  media: MediaStore
  diagnostics: Diagnostic[]
  metadata: PresentationMetadata
}
```

### 4.2 Slide

```ts
interface Slide {
  id: string
  index: number
  name?: string
  background?: Fill
  elements: Element[]
  layoutElements: Element[]
  notes?: string
  transition?: SlideTransition
  timeline?: AnimationTimeline
  diagnostics: Diagnostic[]
}
```

### 4.3 Element

```ts
type Element =
  | TextElement
  | ShapeElement
  | ImageElement
  | TableElement
  | ChartElement
  | GroupElement
  | VideoElement
  | AudioElement
  | ConnectorElement
  | UnknownElement
```

```ts
interface BaseElement {
  id: string
  type: ElementType
  name?: string
  transform: Transform
  opacity: number
  visible: boolean
  zIndex: number
  link?: Hyperlink
  effects?: VisualEffect[]
  source: ElementSource
  diagnostics?: Diagnostic[]
}
```

### 4.4 Transform

```ts
interface Transform {
  x: number
  y: number
  width: number
  height: number
  rotation: number
  flipH: boolean
  flipV: boolean
  matrix: Matrix
}
```

### 4.5 TextElement

```ts
interface TextElement extends BaseElement {
  type: 'text'
  textBody: TextBody
  fill?: Fill
  line?: Line
  verticalAlign?: VerticalAlign
  autoFit?: AutoFit
}

interface TextBody {
  paragraphs: Paragraph[]
  inset?: Insets
  wrap: boolean
  verticalText?: boolean
}

interface Paragraph {
  runs: TextRun[]
  style: ParagraphStyle
  bullet?: Bullet
}

interface TextRun {
  text: string
  style: TextStyle
  link?: Hyperlink
}
```

### 4.6 ShapeElement

```ts
interface ShapeElement extends BaseElement {
  type: 'shape'
  preset?: string
  geometry: Geometry
  fill?: Fill
  line?: Line
  textBody?: TextBody
}
```

### 4.7 ImageElement

```ts
interface ImageElement extends BaseElement {
  type: 'image'
  mediaId: string
  crop?: CropRect
  filters?: ImageFilter[]
  clipGeometry?: Geometry
}
```

### 4.8 TableElement

```ts
interface TableElement extends BaseElement {
  type: 'table'
  rows: TableRow[]
  columns: TableColumn[]
  cells: TableCell[][]
  style?: TableStyle
}
```

### 4.9 AnimationTimeline

```ts
interface AnimationTimeline {
  steps: PlaybackStep[]
  raw?: unknown
  diagnostics: Diagnostic[]
}

interface PlaybackStep {
  id: string
  trigger: 'onClick' | 'withPrevious' | 'afterPrevious'
  effects: AnimationEffect[]
  duration: number
}
```

## 5. 样式系统

企业级 PPT 播放器的关键是样式级联。

PPT 样式来源：

```text
theme
master
layout
slide
shape
paragraph
run
```

需要统一 Style Resolver。

目录建议：

```text
src/style/
  ThemeResolver.ts
  ColorResolver.ts
  FontResolver.ts
  TextStyleResolver.ts
  ParagraphStyleResolver.ts
  FillResolver.ts
  LineResolver.ts
  TableStyleResolver.ts
```

核心接口：

```ts
class StyleResolver {
  resolveTextStyle(ctx: TextStyleContext): ResolvedTextStyle
  resolveParagraphStyle(ctx: ParagraphStyleContext): ResolvedParagraphStyle
  resolveFill(ctx: FillContext): Fill | null
  resolveLine(ctx: LineContext): Line | null
}
```

## 6. 布局引擎

### 6.1 Text Layout Engine

这是最难也是最重要的模块。

V1 使用 HTML/SVG 基础文本渲染。V2 使用 Canvas `measureText()` 做自研换行。V3 支持更完整排版：

- 字体 fallback。
- bullet 缩进。
- numbered list。
- line spacing。
- paragraph spacing。
- autoFit。
- vertical text。
- CJK 换行。
- text box inset。
- baseline。
- superscript/subscript。

目录建议：

```text
src/layout/text/
  TextLayoutEngine.ts
  LineBreaker.ts
  FontMetrics.ts
  BulletLayout.ts
  AutoFit.ts
  CjkLineBreaker.ts
```

### 6.2 Group Layout

处理组合元素坐标变换。

```text
group xfrm
  ├── off
  ├── ext
  ├── chOff
  ├── chExt
  ├── rotation
  ├── flip
```

所有子元素最终归一化到 slide 坐标。

```text
src/layout/group/
  GroupTransform.ts
  Matrix.ts
```

### 6.3 Table Layout

```text
src/layout/table/
  TableLayoutEngine.ts
  CellMerge.ts
  BorderConflictResolver.ts
  TableTextLayout.ts
```

重点：

- 行高。
- 列宽。
- 合并单元格。
- 单元格内边距。
- 单元格文字布局。
- 边框优先级。

### 6.4 Chart Layout

第一阶段可以弱化为数据抽取 + 简单渲染。

```text
src/layout/chart/
  ChartLayoutEngine.ts
  AxisLayout.ts
  LegendLayout.ts
  SeriesLayout.ts
```

## 7. 渲染层设计

### 7.1 渲染路线

阶段 1：SVG Renderer。

优点：

- 形状渲染方便。
- 缩放清晰。
- 方便调试。
- CSS 动画容易接入。

阶段 2：Canvas Static Cache。

将静态页缓存到 Canvas / Bitmap，提高播放性能。

阶段 3：Canvas/WebGL Renderer。

适合大文件和复杂动画。

### 7.2 Renderer 接口

```ts
interface Renderer {
  renderSlide(slide: Slide, container: HTMLElement): Promise<void>
  renderElement(element: Element, ctx: RenderContext): Promise<void>
  clear(): void
  resize(width: number, height: number): void
}
```

### 7.3 目录建议

```text
src/render/
  RenderTree.ts
  RenderContext.ts

  svg/
    SvgRenderer.ts
    SvgShapeRenderer.ts
    SvgTextRenderer.ts
    SvgImageRenderer.ts
    SvgTableRenderer.ts

  canvas/
    CanvasRenderer.ts
    CanvasShapeRenderer.ts
    CanvasTextRenderer.ts
    CanvasImageRenderer.ts

  dom/
    DomOverlayRenderer.ts
    LinkOverlay.ts
    SelectionOverlay.ts
```

## 8. 播放运行时

### 8.1 Player Runtime

职责：

- 当前页管理。
- 当前动画 step 管理。
- 键盘控制。
- 鼠标/触控控制。
- 全屏。
- 缩放。
- 预加载。
- 缩略图。
- 演讲者模式。
- 播放状态同步。

目录建议：

```text
src/player/
  PptxPlayer.ts
  PlayerController.ts
  PlayerState.ts
  ViewportManager.ts
  InputController.ts
  KeyboardController.ts
  TouchController.ts
  PreloadManager.ts
  ThumbnailManager.ts
  SpeakerMode.ts
```

### 8.2 Player API

```ts
const player = new PptxPlayer({
  container,
  file,
  mode: 'slideshow',
  renderer: 'svg',
})

await player.load()

player.next()
player.prev()
player.goToSlide(3)
player.enterFullscreen()
player.destroy()
```

### 8.3 事件系统

```ts
player.on('loaded', callback)
player.on('slideChange', callback)
player.on('stepChange', callback)
player.on('error', callback)
player.on('diagnostic', callback)
```

## 9. 动画系统

### 9.1 动画目标

不一开始追求完整 Office 动画，而是支持常见播放体验。

V1 支持页面切换：

- fade。
- push。
- wipe。
- cover。
- uncover。

V2 支持对象进入动画：

- appear。
- fade in。
- fly in。
- wipe。
- zoom。
- split。

V3 支持：

- exit。
- emphasis。
- motion path。
- with previous。
- after previous。
- delay。
- duration。

### 9.2 目录建议

```text
src/animation/
  TimelineParser.ts
  TimelineNormalizer.ts
  TimelinePlayer.ts
  AnimationEffect.ts
  CssAnimationRenderer.ts
  WebAnimationRenderer.ts
  effects/
    fade.ts
    flyIn.ts
    wipe.ts
    zoom.ts
    appear.ts
```

### 9.3 播放模型

```text
slide 1
  step 0: initial state
  step 1: title fade in
  step 2: bullet list appear
  step 3: image fly in
slide 2
  step 0
  step 1
```

用户按下一次右箭头：

```text
如果当前 slide 还有 step -> 播放下一个 step
否则 -> 下一页
```

## 10. Diagnostics 系统

这是企业可控的关键。

企业客户可以接受复杂 SmartArt 暂不支持，但不能接受不知道为什么显示不一样。

### 10.1 Diagnostic 数据结构

```ts
interface Diagnostic {
  code: string
  severity: 'info' | 'warning' | 'error'
  slideIndex?: number
  elementId?: string
  part?: string
  message: string
  detail?: unknown
}
```

### 10.2 常见错误码

```text
XML_PARSE_FAILED
PART_NOT_FOUND
RELATIONSHIP_NOT_FOUND
MEDIA_NOT_FOUND
UNSUPPORTED_SHAPE
UNSUPPORTED_FILL
UNSUPPORTED_EFFECT
UNSUPPORTED_CHART_TYPE
UNSUPPORTED_SMARTART_LAYOUT
UNSUPPORTED_ANIMATION
FONT_SUBSTITUTED
TEXT_LAYOUT_DEGRADED
TABLE_STYLE_DEGRADED
OLE_OBJECT_IGNORED
VML_IGNORED
```

### 10.3 对外输出

```ts
const result = await parser.parse(file)
console.log(result.diagnostics)
```

播放器里也可以提供开发者面板，展示当前页字体替换、不支持效果、SmartArt 降级等信息。

## 11. Worker 与性能架构

### 11.1 线程模型

```text
Main Thread
  - UI
  - Player controls
  - DOM/SVG mount

Parser Worker
  - unzip
  - XML parse
  - OOXML parse
  - model normalize

Layout Worker
  - text layout
  - table layout
  - render tree build

Render Worker 可选
  - OffscreenCanvas
  - bitmap cache
```

### 11.2 加载策略

```text
1. 快速读取 presentation 信息
2. 优先解析第一页
3. 播放器先显示第一页
4. 后台解析后续页
5. 当前页前后各预加载 1～2 页
6. 缩略图低优先级生成
```

### 11.3 缓存

```text
XmlCache
RelationshipCache
MediaCache
SlideModelCache
RenderTreeCache
ThumbnailCache
BitmapCache
FontMetricsCache
```

## 12. 安全与企业控制

纯前端播放器也要考虑企业场景。

### 12.1 权限

- 禁止下载原文件。
- 禁止复制文本。
- 禁止右键。
- 加水印。
- 过期访问。
- 访问审计。

注意：前端防护不能绝对防截图，但可以满足企业合规基础要求。

### 12.2 水印

水印层独立：

```text
Slide Renderer
  +
Watermark Overlay
  +
Annotation Overlay
```

水印内容：

```text
用户姓名 / 工号 / 时间 / IP / 企业名称
```

## 13. 包设计

建议拆包，避免所有东西塞一个库。

```text
@company/pptx-package
@company/pptx-ooxml
@company/pptx-model
@company/pptx-layout
@company/pptx-renderer-svg
@company/pptx-renderer-canvas
@company/pptx-player
@company/pptx-react
@company/pptx-vue
```

如果一开始不想 monorepo，也至少内部按这些模块边界组织。

## 14. 推荐目录结构

```text
packages/
  core/
    src/
      package/
      xml/
      ooxml/
      model/
      style/
      diagnostics/

  layout/
    src/
      text/
      table/
      group/
      chart/

  renderer-svg/
    src/
      SvgRenderer.ts
      renderers/

  renderer-canvas/
    src/
      CanvasRenderer.ts
      renderers/

  player/
    src/
      PptxPlayer.ts
      PlayerController.ts
      AnimationController.ts
      PreloadManager.ts
      ViewportManager.ts

  react/
    src/
      PptxPlayerView.tsx

  demo/
    src/

  fixtures/
    office/
    wps/
    keynote/
    edge-cases/

  tests/
    golden/
    visual/
    performance/
```

## 15. 企业级能力矩阵

建议从第一天就维护这个矩阵。

| 能力 | V1 | V2 | V3 | V4 |
|---|---:|---:|---:|---:|
| PPTX 解包 | 支持 | 支持 | 支持 | 支持 |
| slide/layout/master | 支持 | 支持 | 支持 | 支持 |
| 主题色 | 部分 | 支持 | 支持 | 支持 |
| 文本基础样式 | 部分 | 支持 | 支持 | 支持 |
| 中文字体 fallback | 部分 | 支持 | 支持 | 支持 |
| bullet/numbering | 部分 | 支持 | 支持 | 支持 |
| autoFit | 不支持 | 部分 | 支持 | 支持 |
| 图片 | 支持 | 支持 | 支持 | 支持 |
| 图片裁剪 | 部分 | 支持 | 支持 | 支持 |
| 基础形状 | 支持 | 支持 | 支持 | 支持 |
| 复杂 shape path | 部分 | 部分 | 支持 | 支持 |
| 表格 | 部分 | 支持 | 支持 | 支持 |
| 图表 | 降级 | 部分 | 支持常见 | 增强 |
| SmartArt | 降级 | 降级 | 部分 | 增强 |
| 页面切换 | 部分 | 支持 | 支持 | 支持 |
| 对象动画 | 不支持 | 部分 | 常见支持 | 增强 |
| 视频/音频 | 部分 | 部分 | 支持 | 支持 |
| 备注 | 支持 | 支持 | 支持 | 支持 |
| diagnostics | 支持 | 支持 | 支持 | 支持 |
| Worker 解析 | 支持 | 支持 | 支持 | 支持 |
| 截图回归 | 部分 | 支持 | 支持 | 支持 |

## 16. 技术选型

### 16.1 应用与工程栈

- 应用框架：Vue 3。
- 构建工具：Vite。
- 开发语言：TypeScript。
- 状态管理：Pinia。
- 路由：Vue Router。
- 测试框架：Vitest。
- 代码质量：ESLint + Oxlint + Prettier。

当前仓库已基于 Vue 3 + Vite + TypeScript 初始化，后续 demo、播放器壳、诊断面板、样本页均沿用该技术栈，不额外引入 React 或其他 UI 框架。

### 16.2 PPTX 解包

Phase 0 使用 JSZip 读取 `.pptx` ZIP 包结构。

选择原因：

- 浏览器兼容成熟。
- API 简单，适合快速打通技术验证。
- 足够覆盖早期的 package part、XML、media 读取需求。

后续如果大文件性能成为瓶颈，再评估 fflate 或 zip.js。

### 16.3 XML 解析

XML 解析使用 txml，并在其上封装统一 XML 访问层。

目录建议：

```text
src/xml/
  XmlParser.ts
  XmlNode.ts
  XmlQuery.ts
  namespaces.ts
```

设计要求：

- OOXML 解析层不直接依赖 txml 原始返回结构。
- `XmlParser` 负责调用 txml 并转换为统一节点结构。
- `XmlQuery` 提供 `child`、`children`、`attr`、`text`、`path` 等稳定查询 API。
- namespace、属性、文本节点和节点顺序的兼容处理集中在 XML 层。

这样后续如果需要替换 XML 底层实现或增强 source location，只影响 XML adapter，不影响 OOXML Parser、Normalized Model 和 Renderer。

### 16.4 数据模型

Normalized Model 自研，不复用第三方 PPTX 库的内部结构。

核心原则：

- Renderer、Player Runtime、Layout Engine 只能消费标准化模型。
- OOXML Parser 负责从 PresentationML / DrawingML 等 XML 转换到标准化模型。
- 不支持内容使用 `UnknownElement` 和 `Diagnostic` 表达，不因复杂对象导致整页崩溃。

### 16.5 渲染路线

首阶段使用 SVG Renderer。

选择原因：

- 对 shape、text、image、transform 表达友好。
- 缩放清晰。
- 易调试。
- 后续接入 CSS 动画或 Web Animations API 成本较低。

Canvas 暂不作为首发渲染器，后续用于静态缓存、大文件优化和动画性能增强。

### 16.6 Worker 策略

Phase 0 暂不引入 Worker，先打通解析、模型和渲染闭环。

但 parser API 必须保持异步形式：

```ts
const presentation = await parser.parse(file)
```

后续 Phase 1/Phase 2 再将 Package Reader、XML parse、OOXML parse、model normalize 迁移到 Parser Worker。

### 16.7 初始依赖

Phase 0 初始新增依赖控制为：

```sh
pnpm add jszip txml
```

其余能力优先自研或基于浏览器原生能力封装，避免早期依赖膨胀。

## 17. 计划列表

### Phase 0：技术验证，2～3 周

目标：证明纯 JS 路线可行。

任务：

1. 搭建 TypeScript monorepo。
2. 实现 PPTX unzip。
3. 实现 ContentTypes 解析。
4. 实现 relationship resolver。
5. 读取 presentation 尺寸。
6. 读取 slides 列表。
7. 渲染第一页背景色。
8. 渲染基础文本框。
9. 渲染基础图片。
10. 输出 diagnostics。
11. 做一个最小播放器 demo。

验收：

- 能打开至少 10 个简单 PPTX。
- 能显示第一页。
- 能翻页。
- 解析失败不崩溃。
- 有 diagnostics 输出。

### Phase 1：静态播放 MVP，1～2 个月

目标：能作为内部预览播放器试用。

解析能力：

1. slide/layout/master 解析。
2. theme 解析。
3. 背景解析。
4. 形状解析：rect、roundRect、ellipse、line、arrow。
5. 图片解析。
6. group transform。
7. 文本解析：paragraph、run、font size、font family、color、bold、italic、underline、alignment。
8. 简单表格解析。
9. 备注解析。
10. 页面切换基础解析。

渲染能力：

1. SVG stage。
2. 文本渲染。
3. 图片渲染。
4. shape 渲染。
5. group 渲染。
6. 基础表格渲染。
7. 页面缩放。
8. 缩略图生成。

播放能力：

1. 上一页 / 下一页。
2. 指定页跳转。
3. 键盘控制。
4. 全屏。
5. 适应窗口。
6. 缩略图列表。
7. 加载进度。
8. 错误页面。

验收：

- 100 个简单/中等复杂企业 PPTX，打开成功率 ≥ 90%。
- 常见文本、图片、形状位置基本正确。
- 不支持对象有 warnings。
- 首屏时间可接受。

### Phase 2：企业常用场景增强，2～3 个月

目标：企业汇报类 PPT 可用。

文本增强：

1. 字体级联。
2. theme font。
3. 中文字体 fallback。
4. bullet。
5. numbering。
6. 段前段后。
7. 行距。
8. 缩进。
9. text box inset。
10. autoFit 初版。
11. 垂直对齐。
12. 超链接。

表格增强：

1. 合并单元格。
2. 单元格背景。
3. 单元格边框。
4. 单元格内边距。
5. 单元格文本布局。
6. 表格主题样式。
7. 行高/列宽。
8. 表格文字对齐。

图片与形状增强：

1. 图片裁剪。
2. 图片透明度。
3. 图片滤镜基础。
4. 渐变填充。
5. 图案填充。
6. 线条箭头。
7. 阴影基础。
8. 更多 preset shape。

工程增强：

1. Parser Worker。
2. Slide lazy parse。
3. Media cache。
4. Render tree cache。
5. 预加载前后页。
6. 文件级 diagnostics report。

验收：

- 300 个企业样本打开成功率 ≥ 95%。
- 文本、表格、图片类页面达到可用级。
- 复杂对象可降级且可诊断。
- 大文件不会阻塞主线程。

### Phase 3：基础动画播放，2 个月

目标：从预览升级到播放。

页面切换：

1. fade。
2. push。
3. wipe。
4. cover。
5. uncover。
6. split。
7. zoom。

对象动画：

1. appear。
2. fade in。
3. fly in。
4. wipe。
5. zoom。
6. split。

时间线：

1. 解析 `p:timing`。
2. 标准化 AnimationTimeline。
3. click step。
4. withPrevious。
5. afterPrevious。
6. duration。
7. delay。

播放器增强：

1. stepIndex 状态。
2. 点击播放下一步。
3. 键盘播放下一步。
4. 动画重置。
5. 跳页时恢复目标页状态。
6. 演讲模式基础。

验收：

- 常见 entrance 动画可播放。
- 不支持动画有降级。
- 翻页与动画 step 逻辑稳定。
- 页面切换体验顺滑。

### Phase 4：图表、SmartArt、媒体增强，3～4 个月

目标：覆盖更多真实企业材料。

图表：

1. bar。
2. column。
3. line。
4. pie。
5. area。
6. combo 初版。
7. legend。
8. axis。
9. data label。
10. theme color。

SmartArt：

1. 文本结构抽取。
2. process 类布局。
3. hierarchy 类布局。
4. cycle 类布局。
5. list 类布局。
6. 复杂布局降级。

媒体：

1. video element。
2. audio element。
3. poster。
4. 播放控制。
5. 页面切换时暂停。
6. 全屏播放兼容。

公式：

1. OMath 基础解析。
2. 常见分式。
3. 根号。
4. 上下标。
5. 矩阵降级。
6. LaTeX/MathML 渲染策略。

验收：

- 图表页不再大面积空白。
- SmartArt 至少可读。
- 音视频元素可播放或可诊断。
- 复杂对象不会破坏整页播放。

### Phase 5：企业级稳定性，持续建设

目标：产品化和商业化。

回归体系：

1. Golden JSON。
2. SVG snapshot。
3. Screenshot diff。
4. Performance benchmark。
5. Memory benchmark。
6. Browser compatibility test。
7. Office/WPS/Keynote fixture 分类。

企业能力：

1. 动态水印。
2. 禁止下载配置。
3. 禁止复制配置。
4. 访问审计事件。
5. 分享链接过期。
6. 权限 hook。
7. 插件机制。
8. 私有字体包管理。
9. diagnostics UI。
10. feature matrix 页面。

性能：

1. OffscreenCanvas 预渲染。
2. Canvas static layer。
3. SVG animated layer。
4. 大文件分片处理。
5. Worker pool。
6. 缩略图异步生成。
7. LRU cache。

验收：

- 1000+ 样本持续回归。
- 企业 PPT 打开成功率 ≥ 98%。
- 核心页面渲染稳定。
- 所有不支持能力可追踪。
- 有明确版本兼容报告。

## 17. 里程碑版本规划

### v0.1 Prototype

- 能打开 PPTX。
- 能读取 slide。
- 能显示文本、图片、矩形。
- SVG 渲染。
- 简单翻页。

### v0.2 MVP

- 支持 slide/layout/master。
- 支持基础样式。
- 支持图片、形状、文本、简单表格。
- 支持缩略图、全屏、键盘。
- 有 diagnostics。

### v0.3 Enterprise Preview

- Worker 解析。
- 文本增强。
- 表格增强。
- 图片裁剪。
- group transform。
- 水印。
- 权限控制接口。

### v0.4 Playback

- 页面切换。
- 基础对象动画。
- step 播放。
- 演讲模式基础。

### v0.5 Advanced Elements

- 常见图表。
- SmartArt 降级/部分布局。
- 音视频。
- 公式。

### v1.0 Enterprise Stable

- 样本库回归。
- 截图 diff。
- 性能基准。
- 插件化。
- 完整 diagnostics。
- 企业部署文档。

## 18. 团队配置建议

- 无团队配置建议

## 19. 风险清单

高风险：

1. 文本排版不准。
2. WPS 文件差异大。
3. SmartArt 复杂。
4. 图表复杂。
5. 动画时间线复杂。
6. 大文件性能。
7. 字体缺失。
8. 截图回归成本高。
9. 浏览器差异。
10. 用户期望接近 Office。

应对策略：

1. 明确能力矩阵。
2. 所有不支持能力 diagnostics。
3. 先覆盖企业高频场景。
4. 不承诺 100% Office 一致。
5. 建样本库持续回归。
6. 允许高级对象降级。
7. 优先做“可读、可播、可解释”。

## 20. 最终推荐路线

如果坚持纯 JS 自研，建议战略为：

```text
第一目标：企业静态预览稳定可用
第二目标：常见 PPT 播放体验可用
第三目标：高频对象高保真
第四目标：高级对象可诊断降级
第五目标：持续通过样本库逼近 Office/WPS
```

最重要的前三件事：

1. 先搭好 Package / OOXML / Model / Diagnostics 四层地基。
2. 不要让播放器直接依赖 XML。
3. 从第一天开始建设样本库和能力矩阵。

这样做出来的不是一个临时 PPT 解析工具，而是一个真正可以长期演进的企业级 PPTX Web Player。
