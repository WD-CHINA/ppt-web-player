# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Repository status

This is an early Vue 3 + Vite + TypeScript application for an enterprise PPTX Web Player. The current source tree is still close to the Vite starter template, while the product architecture is documented in `PPTX_WEB_PLAYER_ARCHITECTURE.md`.

Primary product/architecture context:

- `PPTX_WEB_PLAYER_ARCHITECTURE.md` — enterprise PPTX Web Player architecture, module boundaries, capability roadmap, diagnostics strategy, worker/performance direction, and risks.
- `README.md` — Vite/Vue starter instructions and basic command reference.

## Commands

Package manager: `pnpm`.

Common commands:

```sh
pnpm install        # install dependencies
pnpm dev            # start Vite dev server
pnpm build          # run vue-tsc project build and Vite production build
pnpm build-only     # Vite production build only
pnpm type-check     # TypeScript/Vue type checking via vue-tsc --build
pnpm test:unit      # run Vitest unit tests
pnpm lint           # run oxlint and ESLint with autofix
pnpm lint:oxlint    # run oxlint with autofix
pnpm lint:eslint    # run ESLint with autofix and cache
pnpm format         # format src/ with Prettier
pnpm preview        # preview production build locally
```

Run a single unit test file:

```sh
pnpm test:unit src/__tests__/App.spec.ts
```

Run Vitest in watch/filter mode by passing Vitest arguments after the script name, for example:

```sh
pnpm test:unit -- --watch
pnpm test:unit -- -t "mounts renders properly"
```

## Current application structure

The implemented code currently has a minimal Vue app shape:

- `src/main.ts` creates the Vue app, installs Pinia and Vue Router, and mounts `App.vue` to `#app`.
- `src/App.vue` is still starter-template content.
- `src/router/index.ts` creates a router with `createWebHistory(import.meta.env.BASE_URL)` and no routes yet.
- `src/stores/counter.ts` is a starter Pinia setup store example.
- `src/__tests__/App.spec.ts` is a starter Vitest + Vue Test Utils mount test.

Configuration:

- `vite.config.ts` enables Vue, Vue JSX, Vue DevTools, and `@` as an alias for `src`.
- `vitest.config.ts` merges the Vite config and uses `jsdom` for tests.
- `eslint.config.ts` uses Vue essential rules, recommended Vue TypeScript config, Vitest rules for `src/**/__tests__/*`, Oxlint integration, and Prettier compatibility.
- TypeScript project references are split across `tsconfig.app.json`, `tsconfig.node.json`, and `tsconfig.vitest.json`.

## Intended product architecture

Keep future implementation aligned with the architecture document. The intended layers are:

```text
Player UI
Player Runtime
Renderer Layer
Layout Engine
Normalized Model
OOXML Parser
Package Reader
```

Important dependency rule: player/runtime and renderers should consume normalized models, not raw XML.

### Package and XML responsibilities

`Package Reader` should own `.pptx` ZIP/package concerns:

- `[Content_Types].xml`
- `_rels/.rels`
- slide/layout/master/theme/media relationships
- relative path resolution
- XML and media caches
- missing or invalid resource diagnostics

Centralize relationship/path resolution. Avoid scattered string rewrites such as `target.replace('../', 'ppt/')`; route all relationship target resolution through a dedicated resolver.

The XML layer should parse XML into a safe queryable structure while preserving namespaces, node order, attributes, text, and source-location information for diagnostics.

### OOXML parsing and normalized model

`OOXML Parser` converts PresentationML / DrawingML / ChartML / animation XML into normalized models. It should not render.

The normalized model should center on:

- `Presentation`
- `Slide`
- `Element` union: text, shape, image, table, chart, group, video, audio, connector, unknown
- `Transform`
- text structures: `TextBody`, `Paragraph`, `TextRun`
- `AnimationTimeline`
- `Diagnostic`

Unknown or unsupported content should be represented and diagnosed rather than crashing parsing or rendering.

### Style and layout

Style resolution is a core subsystem. PPT style precedence is:

```text
theme → master → layout → slide → shape → paragraph → run
```

Use resolver layers for theme colors, fonts, text style, paragraph style, fill, line, and table style rather than resolving these ad hoc in renderers.

Layout should be separated from rendering, especially for:

- text layout, including CJK line breaking, bullets, numbering, line spacing, insets, autoFit, vertical text, fallback fonts
- group transforms, normalizing children to slide coordinates
- table layout, including row/column sizing, merges, borders, cell insets, and cell text
- chart layout, initially allowed to degrade to data extraction and simple rendering

### Rendering and playback

Initial rendering direction is SVG first, then static Canvas/bitmap caching, then optional Canvas/WebGL for large or animation-heavy decks.

Player runtime owns slide index, animation step state, keyboard/mouse/touch control, fullscreen, viewport scaling, preloading, thumbnails, and speaker-mode behavior.

Animation should start with common transitions and entrance effects. The playback model is click-step based: advance animation steps on a slide before moving to the next slide.

### Diagnostics

Diagnostics are a first-class product requirement. Complex or unsupported PPTX features may degrade, but the system should explain what degraded and where.

Use structured diagnostics with fields equivalent to:

- `code`
- `severity`
- `slideIndex`
- `elementId`
- `part`
- `message`
- `detail`

Common diagnostic categories include XML parse failures, missing parts or relationships, missing media, unsupported shapes/fills/effects/charts/SmartArt/animations, font substitution, degraded text layout, degraded table style, ignored OLE/VML.

## Planned package/module boundaries

The architecture document recommends package boundaries similar to:

```text
packages/core        # package, xml, ooxml, model, style, diagnostics
packages/layout      # text, table, group, chart layout
packages/renderer-svg
packages/renderer-canvas
packages/player
packages/react
packages/demo
packages/fixtures
packages/tests       # golden, visual, performance
```

This repository is currently a single Vite app, but future implementation should still preserve these internal boundaries. Avoid coupling renderer/UI code directly to raw OOXML/XML parsing details.

## Roadmap context

The planned implementation phases are:

- Phase 0: TypeScript proof of concept; unzip PPTX, parse content types/relationships, read slides, render first slide background/basic text/basic images, diagnostics, minimal demo.
- Phase 1: static playback MVP with slide/layout/master, theme, backgrounds, common shapes, images, group transforms, text basics, simple tables, notes, SVG stage, thumbnails, fullscreen, keyboard, error page.
- Phase 2: enterprise common-scene enhancement: text fidelity, bullets/numbering, table fidelity, image crop/transparency, more shape fills/lines/effects, workers, lazy parsing, caches, diagnostics report.
- Phase 3: basic playback animations and slide transitions.
- Phase 4: charts, SmartArt degradation/partial layout, audio/video, math.
- Phase 5: product stability, golden JSON, SVG snapshots, screenshot diff, performance/memory/browser benchmarks, watermarking, permissions, audit, plugins, private font management, diagnostics UI.

Milestone labels in the architecture document: `v0.1 Prototype`, `v0.2 MVP`, `v0.3 Enterprise Preview`, `v0.4 Playback`, `v0.5 Advanced Elements`, `v1.0 Enterprise Stable`.

## Testing expectations as implementation grows

The architecture calls for continuous regression from the start:

- fixture decks grouped by Office / WPS / Keynote / edge cases
- golden JSON for parser/model output
- SVG snapshot tests
- screenshot diff tests
- performance benchmarks
- memory benchmarks
- browser compatibility tests

Add exact commands here when those test suites are introduced.

## Known risks to account for

High-risk areas identified by the architecture document:

- text layout fidelity
- WPS file differences
- SmartArt complexity
- chart complexity
- animation timeline complexity
- large-file performance
- missing fonts
- screenshot regression cost
- browser differences
- user expectations approaching Office fidelity

Mitigation strategy: maintain a capability matrix, emit diagnostics for unsupported features, cover high-frequency enterprise cases first, avoid promising 100% Office equivalence, build a sample library, allow advanced objects to degrade safely.

## Vue and TypeScript coding rules

The Cursor rule `.cursor/rules/vuejs-typescript-best-practices.mdc` has been converted into Claude-readable project guidance. Apply these rules when editing Vue/TypeScript code in this repository:

- Use concise, maintainable, technically accurate TypeScript.
- Prefer functional and declarative patterns; avoid classes.
- Keep code modular and avoid unnecessary duplication.
- Use descriptive variable names with auxiliary verbs where useful, such as `isLoading` and `hasError`.
- Organize files so each file contains closely related content only: exported components, subcomponents, helpers, static content, and local types.
- Use TypeScript for all new code. Prefer `interface` over `type` when defining extensible object shapes.
- Avoid `enum`; prefer typed maps or literal unions for better type safety and flexibility.
- Use Vue Composition API with `<script setup lang="ts">` for Vue components.
- Use `function` declarations for pure functions when it improves hoisting and readability.
- Use `Suspense` around async components with a fallback UI.
- Dynamically import non-critical components.
- For Vite builds, prefer code splitting and optimized chunking when bundle size becomes relevant.
