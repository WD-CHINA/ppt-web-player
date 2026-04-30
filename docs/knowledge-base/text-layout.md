# Text Layout / OOXML Text 解析记录

## TextBody 是段落集合，不是扁平字符串

### 现象

如果只关心 `shape.text` 这种扁平字符串，后续要支持列表、字号、行距、制表符、字段文本时会很快遇到信息丢失。

### 原因

OOXML 文本结构天然是分层的：

- `p:txBody` -> 文本容器
- `a:p` -> paragraph
- `a:r` / `a:fld` / `a:br` / `a:tab` -> run 或特殊文本节点

### 解决方案

核心模型保留 `TextBody -> Paragraph[] -> TextRun[]` 结构，同时额外提供聚合后的 `text` 便于简单消费。

### 关联文件

- `packages/core/src/model/Presentation.ts`
- `packages/core/src/ooxml/presentation/parseTextBody.ts`
- `packages/core/src/ooxml/presentation/parseShape.ts`

## 段落文本与 run 文本要同时保留

### 现象

渲染层和搜索层通常想直接拿整段文本；但样式层又需要 run 粒度的信息。

### 原因

只保留一种视图会让另一类消费场景变复杂：

- 只保留 runs：调用方要自己拼 paragraph 文本
- 只保留 paragraph 文本：样式信息会丢失

### 解决方案

`Paragraph` 同时保留：

- `runs`
- `text`
- `style`

其中 `text` 由 `runs.map((run) => run.text).join('')` 聚合而来。

### 关联文件

- `packages/core/src/model/Presentation.ts`
- `packages/core/src/ooxml/presentation/parseTextBody.ts`

## default run style 是继承基线，不能覆盖 paragraph style 语义

### 现象

处理 `a:defRPr` 时，容易把它当成段落全部样式，或者误塞进 paragraph 顶层字段。

### 原因

`a:defRPr` 的语义是“paragraph 内未单独声明 run 样式时的默认文本样式”，它服务于 run 继承，不等于 paragraph 自身的布局样式。

### 解决方案

把 `a:defRPr` 解析到 `ParagraphStyle.defaultRunStyle`，run 节点再通过 `mergeTextStyle(defaultRunStyle, runStyle)` 继承并覆盖。

### 关联文件

- `packages/core/src/model/Presentation.ts`
- `packages/core/src/ooxml/presentation/parseTextBody.ts`

## 特殊文本节点不能丢

### 现象

如果只解析 `a:r`，会漏掉 tab、line break、field text，最终 paragraph 文本和 PowerPoint 中看到的不一致。

### 原因

OOXML 文本里除了普通 run，还常见：

- `a:fld`：字段文本，例如 slide number
- `a:br`：换行
- `a:tab`：制表符
- 直接出现的 `a:t`

### 解决方案

`collectRuns()` 必须显式处理这些节点，并将它们映射成统一 `TextRun` 序列：

- `a:fld` -> run
- `a:br` -> `{ text: '\n' }`
- `a:tab` -> `{ text: '\t' }`
- `a:t` -> 纯文本 run

### 关联文件

- `packages/core/src/ooxml/presentation/parseTextBody.ts`
- `packages/core/src/__tests__/parsePptx.spec.ts`

## paragraph style 和 text style 要严格分层

### 现象

缩进、对齐、行距、bullet、rtl 与 bold、italic、fontSize、fontFace 这两类字段很容易混在一起。

### 原因

两者都来自 OOXML 文本节点，但作用域不同：

- paragraph style 作用于整段布局和列表语义
- text style 作用于 run 级字形与装饰

### 解决方案

保持两层模型分离：

- `ParagraphStyle`：`align`、`level`、`indent`、`marginLeft`、`marginRight`、`defaultTabSize`、`rtl`、`bullet`、`lineSpacing`、`spaceBefore`、`spaceAfter`、`defaultRunStyle`
- `TextStyle`：`bold`、`italic`、`underline`、`fontSize`、`color`、`fontFace`

### 关联文件

- `packages/core/src/model/Presentation.ts`
- `packages/core/src/ooxml/presentation/parseTextBody.ts`

## theme 颜色解析要贯穿 paragraph 与 run 两条链路

### 现象

某些文本颜色能解析出来，某些列表符号颜色却丢失，通常不是颜色算法问题，而是 theme 没沿着整条链路传下去。

### 原因

run、default run style、bullet 都可能依赖 `schemeClr`。如果只有其中一部分拿到 `PresentationTheme`，同一段文本会出现半解析状态。

### 解决方案

确保 `parseTextBody()`、`parseParagraphStyle()`、`parseTextRun()`、`parseFieldRun()`、`parseBulletStyle()` 都能在需要时拿到 theme。

### 关联文件

- `packages/core/src/ooxml/presentation/parseTextBody.ts`
- `packages/core/src/ooxml/drawing/Fill.ts`
- `packages/core/src/__tests__/parsePptx.spec.ts`

## 空文本段落是否保留要明确约定

### 现象

PowerPoint XML 里可能存在无可见文本的段落；若不统一处理，调用方会遇到 paragraph 数量与视觉段落数量不稳定的问题。

### 原因

当前实现会过滤 `paragraph.text.length === 0` 的段落，只保留有聚合文本内容的 paragraph。

### 解决方案

在文档和实现中都明确这一约定：当前 core 以“可消费文本内容”为主，不保留空 paragraph。若未来 renderer 需要精确空行语义，再单独评估模型扩展。

### 关联文件

- `packages/core/src/ooxml/presentation/parseTextBody.ts`

## 文本模型保留 OOXML 语义单位时，renderer 必须显式换算

### 现象

文本解析层测试都通过，但一接到 renderer 就出现字号爆大、段落间距离谱、缩进错位。

### 原因

这是一个典型的“模型正确、单位错用”的跨层问题：

- parser/model 层为了保留 OOXML 语义，会继续使用原始数值
- renderer 层如果把这些数值直接当成 SVG/CSS 单位，就会立刻失真

典型例子：

- `fontSize: 1800` 表示 `18pt`
- `marginLeft: 457200` 表示约 `36px`
- `lineSpacing.percent: 150000` 表示 `1.5x`

### 解决方案

在知识边界上明确两件事：

1. `core` 负责把文本结构和样式语义保留下来，不强行替 renderer 决定浏览器单位
2. `renderer` 负责把这些语义单位映射成当前渲染后端可消费的数值

这意味着文本问题排查时，不仅要看“字段有没有解析出来”，还要看“renderer 有没有先做单位归一化”。

这条约束对 `renderer-svg` 和 `renderer-canvas` 都成立。后者即使不生成 SVG/CSS，也不能直接拿 OOXML 原始数值喂给 canvas API；至少要先完成字号、缩进、段前后距、行距这些基础换算，再谈更高保真的文本测量与换行。

### 关联文件

- `packages/core/src/model/Presentation.ts`
- `packages/core/src/ooxml/presentation/parseTextBody.ts`
- `packages/renderer-svg/src/index.ts`

## layout/master text defaults 只作为 run 继承输入

### 现象

slide 上的 placeholder 文本没有直接声明字号、字体或颜色，但 PowerPoint 里仍能继承 layout/master 上定义的默认文本样式。

### 原因

placeholder 文本样式常落在 layout/master 的 `p:txBody/a:lstStyle/a:lvlNpPr/a:defRPr` 或首段 `a:pPr/a:defRPr` 中。slide 本身只携带 placeholder metadata 和实际文本，如果解析 slide 前没有找到对应 layout/master placeholder，就无法得到默认 run 样式。

### 解决方案

当前 core 使用最小继承规则：

- master/layout 解析阶段提取 placeholder `text` defaults
- slide shape 通过 placeholder `idx` 优先、`type` 次之匹配 layout/master defaults
- paragraph style 按 level 合并 inherited defaults 与 slide 直接样式
- run style 按 default run style -> run direct style 合并

这仍不是完整 Office text style cascade；未覆盖复杂 placeholder fallback、theme font scheme 全链路、完整 text style precedence。

### 关联文件

- `packages/core/src/ooxml/presentation/parsePlaceholderStyle.ts`
- `packages/core/src/ooxml/presentation/parseShape.ts`
- `packages/core/src/ooxml/presentation/parseTextBody.ts`
- `packages/core/src/__tests__/parsePptx.spec.ts`

## 遇到文本解析异常时的排查清单

### 现象

常见表现有：文本缺字、换行没了、默认样式没继承、bullet 样式不见、整段文字被当成普通 shape。

### 原因

问题通常分布在文本容器识别、run 收集、样式继承、theme 颜色映射四层。

### 解决方案

建议按下面顺序排查：

1. 确认 shape 下是否存在 `p:txBody`
2. 检查 `a:p` 是否被完整收集
3. 检查 `collectRuns()` 是否覆盖 `a:r` / `a:fld` / `a:br` / `a:tab` / `a:t`
4. 检查 `defaultRunStyle` 是否正确合并到 run 样式
5. 检查 paragraph style 与 text style 是否被混用
6. 如果涉及 `schemeClr`，确认 theme 是否沿调用链传入
7. 如果文本存在但元素不是 `text`，回看 `parseShapeElement()` 的判定逻辑

### 关联文件

- `packages/core/src/ooxml/presentation/parseTextBody.ts`
- `packages/core/src/ooxml/presentation/parseShape.ts`
- `packages/core/src/__tests__/parsePptx.spec.ts`
