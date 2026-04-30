# ppt-web-player

企业级 PPTX Web Player，当前采用 Vue 3 + Vite + TypeScript + pnpm workspace。

## 前端框架约束

本项目只允许使用 Vue 作为 JavaScript 视图框架。禁止引入 React、Svelte、Solid、Angular 或其他 JavaScript 视图框架。

## 项目交付规范

本项目要求“功能实现”和“文档沉淀”同时完成，二者共同构成完整交付。

也就是说，任何重要变更都不仅要改代码，还要同步判断是否需要：

- 更新 `README.md`、架构文档、模块说明、开发计划等文档
- 更新或新增 `docs/knowledge-base/` 下的知识库记录
- 记录本次实现中的关键经验、坑点、约束和排查方法

适用范围包括但不限于：

- 新功能开发
- bug 修复
- 解析/渲染能力扩展
- 数据模型调整
- 重构
- diagnostics、tests、fixtures 相关变更

验收时，不只看功能是否可用，还要看：

- 是否完成必要验证或测试
- 是否更新相关文档
- 是否把可复用经验沉淀进仓库

详细约束见 `docs/project-conventions.md`、`CONTRIBUTING.md` 和 `docs/change-checklist.md`。

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
