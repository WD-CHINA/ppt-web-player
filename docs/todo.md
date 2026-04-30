# PPTX Web Player 4 周开发计划

基于 `docs/PPTX_WEB_PLAYER_ARCHITECTURE.md` 与 `CLAUDE.md`，当前项目最合适的推进方式不是先堆 UI，也不是同时铺开所有 package，而是严格按 `Package Reader -> OOXML Parser -> Normalized Model -> Layout -> Renderer -> Runtime -> UI` 的路径，先把 `core` 做深，再逐层打通最小可用闭环。

目标是在 4 周内，把项目从当前的 Phase 0 解析验证版，推进到 Phase 1 静态播放 MVP 雏形。

## 4 周总目标

- 打稳 `packages/core`，让 PPTX 解析输出更稳定、模型更完整。
- 建立 `layout -> renderer-svg -> player -> app` 的最小闭环。
- 把 diagnostics、fixtures、tests 纳入主流程。
- 明确哪些先做，哪些后做，避免项目横向发散。

## 第 1 周：夯实 core 数据模型和解析边界

### 目标

把当前“能解析”提升到“结构清晰、后续可扩展”。

### 重点工作

- 重构 `packages/core/src/model/Presentation.ts`。
- 为 `Presentation`、`Slide`、`Element` 建立更稳定的标准化模型。
- 给 `TextElement` 引入更细的文本结构预留，如 `TextBody / Paragraph / Run`。
- 补充 `source`、`part`、`zIndex`、`opacity`、`visible` 等后续 render/layout 必需字段。
- 统一 `UnknownElement` 和 diagnostics 的关联方式。

### 解析层工作

- 拆分 `packages/core/src/ooxml/presentation/parseSlide.ts`。
- 从一个大函数改成多个小解析器：
  - `parseShape`
  - `parseImage`
  - `parseTextBody`
  - `parseGeometry`
- 保持“解析器输出 normalized model，不直接面向 UI”。

### 测试工作

- 为 `parsePptx`、`parsePresentation`、`parseSlide` 增加 fixture 测试。
- 为 diagnostics 增加快照或断言测试。

### 本周产出

- `core` 模型升级完成。
- slide 基础元素解析模块化。
- 单元测试覆盖主要 parsing path。

## 第 2 周：补最小样式继承链和 diagnostics 体系

### 目标

解决“解析出来了，但样式来源不明确”的问题。

### 重点工作

- 在 `packages/core` 增加 `theme/master/layout/slide` 的最小骨架。
- 先不追求完整 Office fidelity，只先打通引用和继承路径。
- 解析：
  - `presentation -> slide master` 引用
  - `slide -> layout` 引用
  - `theme` 基础颜色/字体入口
- 建立样式解析/合并的最小规则。

### diagnostics 工作

- 统一错误码命名规范，例如：
  - `PRESENTATION_NOT_FOUND`
  - `RELATIONSHIP_NOT_FOUND`
  - `UNSUPPORTED_SLIDE_ELEMENT`
  - `STYLE_INHERITANCE_INCOMPLETE`
- 明确 severity 使用规则：`info / warning / error`。
- 每类 diagnostics 尽量附带 `part`、`slideIndex`、`elementId`。

### app 验证工作

- 在 `packages/app/src/App.vue` 中增强 diagnostics 展示。
- 支持按 slide、severity、code 维度查看。

### 本周产出

- 最小样式继承框架已覆盖 theme `schemeClr`、master/layout background、placeholder fill/line/text defaults。
- slide 解析已按 direct > layout > master 合并最小样式，renderer 继续消费合并后的 normalized model。
- diagnostics 面板已支持 severity、code、slide 过滤，并展示 `part`、`slideIndex`、`elementId`、`detail`。
- parser 对 placeholder 样式未匹配等降级情况会输出 `STYLE_INHERITANCE_INCOMPLETE`。

## 第 3 周：做 renderer-svg MVP，打通静态渲染

### 目标

让 normalized model 真正被独立 renderer 消费。

### 原则

不要继续让 `app` 自己硬编码预览逻辑，开始把渲染责任迁到 `packages/renderer-svg`。

### 重点工作

- 在 `packages/renderer-svg/src/index.ts` 附近建立 SVG renderer 入口。
- 定义 renderer 输入接口：消费 `Presentation`、`Slide` 或 layout 结果。
- 优先支持这些元素的静态渲染：
  - slide background
  - text
  - image
  - basic shape
  - connector
- 先支持最常见的 `fill/line/geometry`。

### app 层调整

- `packages/app` 只负责：
  - 文件上传
  - slide 切换
  - 调用 parser
  - 把 active slide 交给 renderer
- 不再把大量 SVG 细节逻辑都放在 `App.vue`。

### 测试工作

- 增加基础渲染验证。
- 至少保证几份公开 sample deck 不报错并有可见输出。

### 本周产出

- `renderer-svg` 最小可用。
- `renderer-canvas` 已具备同源 normalized model 的基础静态渲染能力，用于对照 SVG 后端差异。
- `app` 从“自己画预览”转向“调用 renderer”。
- 形成 `parser -> layout -> renderer` 的正式边界。
- renderer 测试和知识库记录覆盖 SVG/Canvas 的当前能力差异。
- 解析/渲染增量已补齐图片裁剪、图片透明度、文本框 inset、wrap=false、基础 CJK 换行与 normal autoFit 近似缩放。
- 图片裁剪和透明度异常会进入 structured diagnostics，renderer 继续只消费 normalized model。
- 已新增 WPS 参考图视觉回归脚本，可对 SVG/Canvas 预览区域生成 actual、diff 和 JSON report，用于定位大块缺图、层级和样式退化。
- slide background 已从普通 `Fill` 扩展为 `SlideBackground`，core 可解析 `p:bg/p:bgPr/a:blipFill` 图片背景，SVG/Canvas 会在元素前先绘制背景图。

## 第 4 周：补 player 最小运行时，形成 Phase 1 雏形

### 目标

从“解析/渲染 demo”升级到“最小播放器”。

### 重点工作

- 在 `packages/player/src/index.ts` 搭最小 runtime。
- runtime 首批职责只做：
  - active slide index
  - next/prev/first/last
  - keyboard navigation
  - viewport scale
  - fit-to-screen / fit-width
- 先不要做复杂动画调度。

### app 层工作

- 将导航状态从 `App.vue` 的本地逻辑，逐步迁到 `player`。
- 增加基础播放器 UI：
  - 当前页 / 总页数
  - 上一页 / 下一页
  - 缩放
  - 缩略图入口可先占位

### 质量工作

- 跑通 `pnpm build`、`pnpm test:unit`。
- 补一批典型 PPTX fixture。
- 明确“不支持项列表”和降级行为。

### 本周产出

- 最小 Player Runtime。
- 最小静态 PPTX Web Player 雏形。
- Phase 1 MVP 的第一版可演示结果。

## 每周结束时的验收标准

### 第 1 周验收

- `core` 模型比现在更稳定。
- `parseSlide.ts` 不再是单个大解析器。
- 样本 PPTX 解析不回退。
- 本周涉及的模型/解析约束已同步到 `docs/` 或知识库。

### 第 2 周验收

- 样式继承有最小链路。
- diagnostics 面板能解释主要降级原因。
- 新增 parser 行为都有对应测试。
- 本周新增的 diagnostics / style 经验已沉淀到知识库。

### 第 3 周验收

- `renderer-svg` 独立消费 normalized model，并提供正式 renderer 入口。
- `app` 中渲染逻辑明显收缩，只负责解析、翻页和渲染调用。
- 至少 3~5 个样本 PPTX 能静态展示。
- parser -> renderer 边界、渲染约束和排查经验有对应文档记录。

### 第 4 周验收

- 用户能上传 PPTX、翻页、查看静态渲染结果。
- 键盘导航和缩放可用。
- 不支持项会通过 diagnostics 明示。
- 本阶段形成的 runtime / player 约束、降级策略和经验已完成文档沉淀。

## 这 4 周里不要急着做的事

- 不要先做 `renderer-canvas` 主实现。
- 不要先做复杂动画系统。
- 不要先做 `SmartArt`、`Chart`、`Math`、`Audio/Video`。
- 不要把 `app` 做成复杂产品 UI。
- 不要跳过 `layout` 或 `normalized model`，直接让 renderer 读 raw XML。

## 建议的里程碑命名

- Week 1：Parser Foundation
- Week 2：Style and Diagnostics
- Week 3：SVG Render MVP
- Week 4：Player Runtime MVP

## 一句话总结

第 1、2 周做深 `core`，第 3 周打通 `renderer-svg`，第 4 周补 `player`，形成最小静态播放闭环。