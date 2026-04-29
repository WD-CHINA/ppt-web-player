# ppt-web-player

企业级 PPTX Web Player，当前采用 Vue 3 + Vite + TypeScript + pnpm workspace。

## 前端框架约束

本项目只允许使用 Vue 作为 JavaScript 视图框架。禁止引入 React、Svelte、Solid、Angular 或其他 JavaScript 视图框架。

## Workspace 结构

```text
packages/
  core/             # PPTX package、XML、OOXML、model、style、diagnostics 等核心能力
  layout/           # 文本、表格、组合、图表布局能力
  renderer-svg/     # SVG 渲染器
  renderer-canvas/  # Canvas 渲染器
  player/           # 播放运行时
  app/              # Vue 3 + Vite 示例应用
```

## 推荐 IDE

- VS Code
- Vue (Official) 插件
- 禁用 Vetur

## 项目安装

```sh
pnpm install
```

## 开发

```sh
pnpm dev
```

该命令会启动 `packages/app` 的 Vite 开发服务器。

## 类型检查

```sh
pnpm type-check
```

## 构建

```sh
pnpm build
```

## 单元测试

```sh
pnpm test:unit
```

运行 app 单个测试文件示例：

```sh
pnpm --filter @pptx-player/app test:unit src/__tests__/App.spec.ts
```

## Lint

```sh
pnpm lint
```

## 格式化

```sh
pnpm format
```
