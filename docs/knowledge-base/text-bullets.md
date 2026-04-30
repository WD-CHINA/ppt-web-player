# Text / Bullet 样式解析记录

## bullet 颜色节点不能复用普通文本颜色解析

### 现象

为段落 bullet 增加颜色解析后，`a:buClr` 明明存在，但测试里解析结果仍然没有 `color` 字段。

### 原因

普通文本 run 和默认 run 样式的颜色结构通常是：

```xml
<a:rPr>
  <a:solidFill>
    <a:srgbClr val="FF6600"/>
  </a:solidFill>
</a:rPr>
```

而 bullet 颜色节点是：

```xml
<a:buClr>
  <a:schemeClr val="accent4"/>
</a:buClr>
```

`a:buClr` 自身就是颜色容器，不再包一层 `a:solidFill`。如果直接复用 run 的颜色解析逻辑，会查不到子节点。

### 解决方案

为 bullet 单独实现 `parseBulletColor()`：

- 直接读取 `a:buClr > a:srgbClr`
- 或读取 `a:buClr > a:schemeClr`
- `schemeClr` 仍然通过 theme color scheme 做映射

### 关联文件

- `packages/core/src/ooxml/presentation/parseTextBody.ts`
- `packages/core/src/__tests__/parsePptx.spec.ts`

## bullet 样式应挂在 ParagraphStyle.bullet，而不是 TextStyle

### 现象

在扩展列表符号字体、颜色、字号时，容易下意识把它们塞到 `TextStyle`，因为字段形态和 run 样式很像。

### 原因

bullet 样式属于段落级语义，不属于具体文本 run：

- 一个段落可以有 bullet，但内部多个 run 仍然保留各自文本样式
- bullet 的字号、字体、颜色可能与 run 默认样式不同
- auto-number 和 character bullet 也共用同一套段落级配置

### 解决方案

统一挂到 `ParagraphStyle.bullet`，在 `BulletStyle` 上扩展：

- `fontFace`
- `color`
- `fontSize`

这样渲染层在画列表符号时可以直接读取段落 bullet 配置，不会和正文 run 样式混淆。

### 关联文件

- `packages/core/src/model/Presentation.ts`
- `packages/core/src/ooxml/presentation/parseTextBody.ts`

## bullet 的 theme 颜色解析依赖 paragraph 解析链路传入 theme

### 现象

`a:buClr > a:schemeClr` 存在时，若 `parseBulletStyle()` 没拿到 theme，最终 `color` 会丢失。

### 原因

theme color resolution 不是 bullet 私有能力，而是 OOXML 文本样式通用能力。bullet 只是在段落级解析阶段更早发生。

### 解决方案

让 `parseParagraphStyle()` 调用 `parseBulletStyle(properties, theme)`，保持 bullet 与 run/default run style 一样都能拿到 `PresentationTheme`。

### 关联文件

- `packages/core/src/ooxml/presentation/parseTextBody.ts`

## bullet 字号有三种来源，测试必须分别覆盖

### 现象

如果只测 `a:buSzPts`，很容易误以为 bullet 字号问题已经完整解决，但真实 PPTX 还会出现百分比和跟随正文两种形式。

### 原因

OOXML 里 bullet 字号常见有三种节点：

- `a:buSzPts`：显式 point 值
- `a:buSzPct`：相对正文的百分比值
- `a:buSzTx`：跟随文本字号

它们的语义不同，解析器必须至少区分来源。

### 解决方案

当前模型先统一输出为 `BulletStyle.fontSize`：

- `buSzPts` -> 直接返回其数值
- `buSzPct` -> 直接返回其数值
- `buSzTx` -> 当前以 `0` 作为“follow text”的占位值

注意：`buSzTx` 的语义化建模未来仍可继续优化，例如改成单独字段或枚举，而不是使用哨兵值。

### 关联文件

- `packages/core/src/ooxml/presentation/parseTextBody.ts`
- `packages/core/src/__tests__/parsePptx.spec.ts`

## 不要手填 theme 颜色断言，先看测试主题真实映射

### 现象

补测试时，初次手写的 `accent4` / `accent6` 预期颜色与实际解析值不一致，导致单测失败。

### 原因

测试 PPTX 使用的是该 fixture/theme 自己的配色，并不一定等于常见的 Tailwind 色值或人工猜测值。theme color name 只是语义键，不是固定 RGB。

### 解决方案

写 `schemeClr` 相关断言前，先确认当前测试输入里的 theme color scheme 实际映射结果，再回填预期值。

### 关联文件

- `packages/core/src/__tests__/parsePptx.spec.ts`

## 遇到类似列表样式问题时的排查清单

### 现象

列表符号缺字形、颜色不对、字号不对，或者 renderer 端明明读到段落文本却画不出列表符号。

### 原因

这类问题通常横跨模型、OOXML 解析、theme 解析、renderer 四层，不是单点故障。

### 解决方案

建议按下面顺序排查：

1. 确认 XML 里使用的是 `a:buChar`、`a:buAutoNum` 还是 `a:buNone`
2. 检查是否存在 `a:buFont`、`a:buClr`、`a:buSzPts|a:buSzPct|a:buSzTx`
3. 如果是 `schemeClr`，确认 `parseBulletStyle()` 是否拿到了 theme
4. 检查 `BulletStyle` 模型是否已经承载所需字段
5. 确认测试断言使用的是 fixture 真实 theme 映射值
6. 如果解析层无误，再检查 renderer 是否真正消费了 `paragraph.style.bullet`

### 关联文件

- `packages/core/src/model/Presentation.ts`
- `packages/core/src/ooxml/presentation/parseTextBody.ts`
- `packages/core/src/__tests__/parsePptx.spec.ts`
