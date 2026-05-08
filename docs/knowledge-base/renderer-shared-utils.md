# 渲染器共享工具

## 概述

`packages/core/src/renderer/rendererUtils.ts` 集中管理 SVG 和 Canvas 渲染器共用的工具函数。

## 导出的函数

| 函数 | 用途 |
|------|------|
| `orderedVisibleElements` | 按 zIndex 排序并过滤不可见元素 |
| `solidFillColor` | 从 Fill 提取 solid 颜色值 |
| `svgFill` | Fill → SVG fill 颜色字符串 |
| `svgFillOpacity` | Fill → 透明度 |
| `svgStrokeOpacity` | LineStyle → 描边透明度 |
| `svgStrokeDasharray` | LineStyle → SVG stroke-dasharray |
| `canvasLineDash` | LineStyle → Canvas 数值数组 |
| `svgMarkerStart` | 线条起点箭头标记 |
| `svgMarkerEnd` | 线条终点箭头标记 |

## 关联文件

- 定义：`packages/core/src/renderer/rendererUtils.ts`
- 导出：`packages/core/src/index.ts`
- 消费方：`packages/renderer-svg/src/index.ts`、`packages/renderer-canvas/src/index.ts`

## 变更历史

- Week 4：从 renderer-svg 和 renderer-canvas 中提取重复代码，集中到 core

---

# txml.d.ts 多包重复

## 现象

`packages/core/src/txml.d.ts`、`packages/layout/src/txml.d.ts`、`packages/renderer-svg/src/txml.d.ts`、`packages/renderer-canvas/src/txml.d.ts` 四份完全相同的模块声明文件。

## 原因

1. `txml` npm 包的 `package.json` 中 `exports` 字段未包含 `.d.ts` 文件路径
2. TypeScript 的 `moduleResolution: "bundler"` 遵守 `exports` 字段，因此无法解析 `txml` 的类型声明
3. 需要手动 `declare module 'txml'` 提供类型
4. Core 的 `XmlParser.ts` 直接 import txml，而 layout/renderer-* 依赖 core
5. 各包独立运行 `tsc --noEmit` 时，因为 `@pptx-player/core: workspace:*` 解析到源码，layout/renderer-* 的类型检查会拉入 `XmlParser.ts`，从而也需要 txml 声明

## 解决方向（待后续处理）

- 选项 A：Core 封装 txml，只通过自己的类型暴露 XML 层，消费者不再传递依赖 txml
- 选项 B：使用项目级 `typeRoots` + 共享 `types/` 目录
- 选项 C：等待 txml 修复 exports 字段
- 当前做法：保持四份拷贝，每次修改同步更新

## 排查方法

如果在其他包中遇到 `Could not find a declaration file for module 'txml'`：

1. 确认该包是否（传递）依赖了 `@pptx-player/core` 中的 `XmlParser` 或 `PptxPackage`
2. 如果是，在该包的 `src/` 下添加 `txml.d.ts`（与 core 保持一致）
3. 或者：在 `tsconfig.json` 中添加 `"skipLibCheck": true` 可能不够，因为这影响的是 `.d.ts` 文件而非 `.ts` 源文件的 import 解析

## 关联文件

- `packages/core/src/txml.d.ts`
- `packages/layout/src/txml.d.ts`
- `packages/renderer-svg/src/txml.d.ts`
- `packages/renderer-canvas/src/txml.d.ts`
- `packages/core/src/xml/XmlParser.ts`（实际 import txml 的文件）
