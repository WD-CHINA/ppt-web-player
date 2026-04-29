# Phase 0 PPTX Package 解析记录

## txml 类型声明无法通过 package exports 解析

### 现象

`tsc` 报错无法找到 `txml` 的声明文件，虽然包内存在 `dist/index.d.ts`。

### 原因

`txml` 当前 package exports 没有让 TypeScript 在 `moduleResolution: Bundler` 下正确解析类型。

### 解决方案

`XmlParser` 从 `txml` 正式入口导入 `parse` 和 `tNode` 类型，并在 `packages/core/src/txml.d.ts` 增加最小模块声明。OOXML 解析层只消费统一的 `XmlNode` 和 `XmlQuery` API。

注意：`txml@5.2.1` 导出的类型名是 `tNode` 和 `TParseOptions`，不是 `TNode` / `ParseOptions`。

### 关联文件

- `packages/core/src/txml.d.ts`
- `packages/core/src/xml/XmlParser.ts`
- `packages/core/src/xml/XmlQuery.ts`

## XML 声明节点不能作为文档根节点

### 现象

真实 PPTX 的 `presentation.xml` 解析后 slide 列表为空，但页面尺寸能从默认逻辑继续返回。

### 原因

`txml` 会把 `<?xml ...?>` 声明解析成 `?xml` 节点，如果直接取第一个节点，会误把 XML 声明当成根节点。

### 解决方案

`XmlParser` 选择根节点时跳过名称以 `?` 开头的声明节点。

### 关联文件

- `packages/core/src/xml/XmlParser.ts`

## relationship target 必须集中解析

### 现象

PPTX 中 slide、media、layout 等关系的 `Target` 可能是相对路径，例如 `../media/image1.png`。

### 原因

relationship target 是相对于关系源 part 所在目录解析的，不能在各解析器里手写字符串替换。

### 解决方案

所有 relationship 路径都通过 `RelationshipResolver` 解析，先取 source part 目录，再归一化 `.` 和 `..` 片段。

### 关联文件

- `packages/core/src/package/RelationshipResolver.ts`
- `packages/core/src/package/PartName.ts`
- `packages/core/src/package/PptxPackage.ts`
