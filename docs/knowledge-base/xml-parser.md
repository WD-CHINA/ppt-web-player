# XML Parser 解析记录

## txml 输出需要先归一化，再暴露给业务层

### 现象

如果业务解析器直接依赖 `txml` 原始结构，后续替换 XML 库、修补类型声明、调整节点字段时，所有解析逻辑都会一起受影响。

### 原因

第三方 XML 库的节点结构、属性类型和文本节点表示方式通常不稳定：

- 属性值可能是 string / number / boolean 混合
- 文本可能散落在子节点数组中
- 根节点前可能混入 XML 声明或其他非元素节点

### 解决方案

先把 `txml` 输出归一化成项目自己的 `XmlNode`，业务层只依赖 `XmlParser` + `XmlQuery`：

- `XmlParser` 负责第三方结构适配
- `XmlQuery` 负责 `child` / `children` / `attr` / `text` / `path`
- OOXML 解析器不要直接读 `txml` 原始节点

### 关联文件

- `packages/core/src/xml/XmlParser.ts`
- `packages/core/src/xml/XmlQuery.ts`
- `packages/core/src/xml/XmlNode.ts`

## XML 声明节点不能被当成文档根节点

### 现象

解析真实 PPTX XML 时，若直接拿解析结果的第一个节点作为根节点，后续 `presentation`、`slide`、`theme` 等查询可能全部失败。

### 原因

`txml` 会把 `<?xml ...?>` 声明解析成名称以 `?` 开头的节点。这个节点不是业务根节点。

### 解决方案

选择根节点时跳过名称以 `?` 开头的节点，只把真正的元素节点作为返回根节点。

### 关联文件

- `packages/core/src/xml/XmlParser.ts`
- `packages/core/src/__tests__/XmlParser.spec.ts`

## 属性值要统一归一化成 string

### 现象

在解析 OOXML 属性时，如果有的属性是 number / boolean，有的是 string，业务代码里会不断出现类型分支和隐式转换。

### 原因

XML 属性在语义上本就是字符串；数值和布尔值只是后续解析器的解释结果。

### 解决方案

在 `XmlParser` 阶段统一把属性值转成 string。后续所有数值、布尔值都由业务解析器显式调用 `parseNumber()`、`parseBooleanFlag()` 之类的 helper 处理。

### 关联文件

- `packages/core/src/xml/XmlParser.ts`
- `packages/core/src/ooxml/presentation/parseTextBody.ts`

## 文本查询必须递归聚合子节点文本

### 现象

某些节点明明含有文本，但直接读取当前节点自身 `text` 字段会拿不全，尤其在 mixed content 或嵌套结构下更明显。

### 原因

文本内容可能分散在当前节点和后代节点里；只读当前层会漏掉嵌套文本。

### 解决方案

统一通过 `xml.text(node)` 获取文本，由 `XmlQuery.text()` 递归拼接当前节点与所有后代节点文本。

### 关联文件

- `packages/core/src/xml/XmlQuery.ts`
- `packages/core/src/ooxml/presentation/parseTextBody.ts`

## 命名空间前缀要按字面值匹配，不要提前剥离

### 现象

OOXML 查询里如果试图把 `p:sp`、`a:p`、`r:embed` 之类前缀先剥掉，再做“通用标签名”匹配，容易造成错误命中。

### 原因

当前项目的 `XmlNode.name` 保留了原始前缀，解析器查询也是围绕这些完整名称编写的。贸然剥离前缀会破坏已有查询假设。

### 解决方案

保持查询条件与 XML 实际节点名一致，例如：

- `p:txBody`
- `a:p`
- `a:r`
- `a:solidFill`

只有在确实要做命名空间抽象时，才引入新的统一查询层。

### 关联文件

- `packages/core/src/xml/XmlQuery.ts`
- `packages/core/src/ooxml/presentation/parseTextBody.ts`
- `packages/core/src/ooxml/presentation/parseShape.ts`

## 遇到 XML 查询异常时的排查清单

### 现象

表现通常是字段大量 `undefined`、元素数量异常、文本丢失，或者整个 slide/theme/presentation 解析为空。

### 原因

问题常出在根节点选错、节点名写错、属性读取路径不对，或把第三方解析结果的结构假设错了。

### 解决方案

建议按下面顺序排查：

1. 确认 `parseXml()` 返回的根节点名称是否正确
2. 检查是否误把 `?xml` 声明当成根节点
3. 确认查询使用的是完整节点名，例如 `a:r` 而不是 `r`
4. 检查文本是否应该通过 `xml.text()` 递归获取
5. 检查属性是否在当前节点，还是在子节点上
6. 如果改了 XML 库适配层，优先补 `XmlParser.spec.ts`

### 关联文件

- `packages/core/src/xml/XmlParser.ts`
- `packages/core/src/xml/XmlQuery.ts`
- `packages/core/src/__tests__/XmlParser.spec.ts`
