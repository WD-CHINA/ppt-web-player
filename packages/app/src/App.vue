<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, ref, watch } from 'vue'
import { parsePptx, type Diagnostic, type DiagnosticSeverity, type Fill, type Presentation } from '@pptx-player/core'
import { renderSlideToCanvas } from '@pptx-player/renderer-canvas'
import { renderSlideToSvg } from '@pptx-player/renderer-svg'

import { usePlayer } from './composables/usePlayer'
import { sampleDecks } from './samples'

const isLoading = ref(false)
const errorMessage = ref('')
const diagnostics = ref<Diagnostic[]>([])
const presentation = ref<Presentation | null>(null)
const mediaUrls = ref<Record<string, string>>({})
const activeFileName = ref('')
const canvasHost = ref<HTMLCanvasElement | null>(null)
const canvasBitmaps = ref<Record<string, ImageBitmap>>({})

const {
  activeSlideIndex,
  slideCount,
  canGoNext,
  canGoPrevious,
  goToSlide,
  first: goToFirstSlide,
  previous: goToPreviousSlide,
  next: goToNextSlide,
  last: goToLastSlide,
  updateSlideCount,
} = usePlayer({ slideCount: 0 })
const diagnosticSeverityFilter = ref<DiagnosticSeverity | 'all'>('all')
const diagnosticCodeFilter = ref('all')
const diagnosticSlideFilter = ref('all')

const hasResult = computed(() => presentation.value !== null)
const activeSlide = computed(() => presentation.value?.slides[activeSlideIndex.value] ?? presentation.value?.slides[0] ?? null)
const activeSlideSvg = computed(() => {
  if (!presentation.value || !activeSlide.value) {
    return ''
  }

  return renderSlideToSvg({
    presentation: presentation.value,
    slide: activeSlide.value,
    mediaUrls: mediaUrls.value,
    ariaLabel: `${activeSlide.value.part} preview`,
    className: 'slide-preview',
  })
})
const diagnosticCounts = computed<Record<DiagnosticSeverity, number>>(() => ({
  info: diagnostics.value.filter((diagnostic) => diagnostic.severity === 'info').length,
  warning: diagnostics.value.filter((diagnostic) => diagnostic.severity === 'warning').length,
  error: diagnostics.value.filter((diagnostic) => diagnostic.severity === 'error').length,
}))
const diagnosticCodes = computed(() => [...new Set(diagnostics.value.map((diagnostic) => diagnostic.code))].sort())
const diagnosticSlideOptions = computed(() =>
  [
    ...new Set(
      diagnostics.value
        .map((diagnostic) => diagnostic.slideIndex)
        .filter((slideIndex): slideIndex is number => slideIndex !== undefined),
    ),
  ].sort((left, right) => left - right),
)
const filteredDiagnostics = computed(() =>
  diagnostics.value.filter((diagnostic) => {
    const matchesSeverity = diagnosticSeverityFilter.value === 'all' || diagnostic.severity === diagnosticSeverityFilter.value
    const matchesCode = diagnosticCodeFilter.value === 'all' || diagnostic.code === diagnosticCodeFilter.value
    const matchesSlide = diagnosticSlideFilter.value === 'all' || String(diagnostic.slideIndex) === diagnosticSlideFilter.value

    return matchesSeverity && matchesCode && matchesSlide
  }),
)

async function parseInput(input: Blob | ArrayBuffer, fileName: string) {
  isLoading.value = true
  errorMessage.value = ''
  presentation.value = null
  diagnostics.value = []
  activeFileName.value = fileName
  goToSlide(0)
  revokeMediaUrls()
  revokeCanvasBitmaps()

  try {
    const result = await parsePptx(input)
    presentation.value = result.presentation
    diagnostics.value = result.diagnostics
    mediaUrls.value = createMediaUrls(result.media)
    canvasBitmaps.value = await createCanvasBitmaps(result.media)

    if (result.presentation) {
      updateSlideCount(result.presentation.slides.length)
    } else {
      errorMessage.value = '未能解析 presentation.xml，请查看 diagnostics。'
    }
  } catch (error) {
    errorMessage.value = error instanceof Error ? error.message : '解析 PPTX 时发生未知错误。'
  } finally {
    isLoading.value = false
  }
}

async function parseSample(path: string, name: string) {
  const response = await fetch(path)

  if (!response.ok) {
    errorMessage.value = `样本加载失败：${response.status}`
    return
  }

  await parseInput(await response.arrayBuffer(), name)
}

async function handleFileChange(event: Event) {
  const input = event.target as HTMLInputElement
  const file = input.files?.[0]

  if (!file) {
    return
  }

  await parseInput(file, file.name)
  input.value = ''
}

function solidFillColor(fill: Fill | undefined, fallback: string): string {
  return fill?.type === 'solid' ? fill.color : fallback
}

function fillLabel(fill: Fill | undefined): string {
  if (!fill) {
    return 'none'
  }

  return fill.type === 'solid' ? fill.color : 'noFill'
}

function formatDiagnosticDetail(detail: unknown): string {
  return detail === undefined ? '' : JSON.stringify(detail, null, 2)
}

function createMediaUrls(media: Record<string, Blob>): Record<string, string> {
  return Object.fromEntries(Object.entries(media).map(([part, blob]) => [part, URL.createObjectURL(blob)]))
}

async function createCanvasBitmaps(media: Record<string, Blob>): Promise<Record<string, ImageBitmap>> {
  if (typeof createImageBitmap !== 'function') {
    return {}
  }

  const entries = await Promise.all(
    Object.entries(media).map(async ([part, blob]) => {
      try {
        return [part, await createImageBitmap(blob)] as const
      } catch {
        return null
      }
    }),
  )

  return Object.fromEntries(entries.filter((entry): entry is readonly [string, ImageBitmap] => entry !== null))
}

function revokeMediaUrls() {
  for (const url of Object.values(mediaUrls.value)) {
    URL.revokeObjectURL(url)
  }

  mediaUrls.value = {}
}

function revokeCanvasBitmaps() {
  for (const bitmap of Object.values(canvasBitmaps.value)) {
    bitmap.close()
  }

  canvasBitmaps.value = {}
}

function renderActiveSlideToCanvas() {
  if (!presentation.value || !activeSlide.value || !canvasHost.value) {
    return
  }

  const canvas = canvasHost.value
  const context = canvas.getContext('2d')

  if (!context) {
    return
  }

  canvas.width = presentation.value.width
  canvas.height = presentation.value.height

  renderSlideToCanvas({
    presentation: presentation.value,
    slide: activeSlide.value,
    context,
    mediaBitmaps: canvasBitmaps.value,
    clear: true,
  })
}

watch([presentation, activeSlide, canvasBitmaps], async () => {
  await nextTick()
  renderActiveSlideToCanvas()
})

// Keyboard handling is managed by usePlayer composable (attachKeyboard on mount, destroy on unmount).
onBeforeUnmount(() => {
  revokeMediaUrls()
  revokeCanvasBitmaps()
})
</script>

<template>
  <main class="app-shell">
    <section class="hero-panel">
      <p class="eyebrow">PPTX Web Player · Week 4</p>
      <h1>企业可控 PPTX Web Player</h1>
      <p class="description">
        当前阶段已打通 PPTX 解析、双渲染器静态渲染与 Player Runtime MVP。导航与键盘控制已迁入独立 `@pptx-player/player` 包，`app` 仅负责上传、解析调用与 UI 渲染。
      </p>
    </section>

    <section class="control-panel" aria-labelledby="parse-title">
      <h2 id="parse-title">解析 PPTX</h2>

      <div class="actions">
        <label class="upload-card">
          <span>上传本地 PPTX</span>
          <input type="file" accept=".pptx" :disabled="isLoading" @change="handleFileChange" />
        </label>

        <div class="sample-list" aria-label="样本文件">
          <button
            v-for="sample in sampleDecks"
            :key="sample.path"
            type="button"
            :disabled="isLoading"
            @click="parseSample(sample.path, sample.name)"
          >
            {{ sample.name }}
          </button>
        </div>
      </div>
    </section>

    <section class="result-panel" aria-live="polite">
      <div v-if="isLoading" class="status-card">正在解析 {{ activeFileName }}...</div>

      <div v-else-if="errorMessage" class="status-card error-card">
        {{ errorMessage }}
      </div>

      <div v-else-if="hasResult && presentation" class="presentation-card">
        <div>
          <p class="eyebrow">解析结果</p>
          <h2>{{ activeFileName }}</h2>
        </div>

        <dl class="metrics-grid">
          <div>
            <dt>画布宽度</dt>
            <dd>{{ presentation.width }} px</dd>
          </div>
          <div>
            <dt>画布高度</dt>
            <dd>{{ presentation.height }} px</dd>
          </div>
          <div>
            <dt>幻灯片数量</dt>
            <dd>{{ presentation.slides.length }}</dd>
          </div>
          <div>
            <dt>Presentation Part</dt>
            <dd>{{ presentation.metadata.presentationPart }}</dd>
          </div>
        </dl>

        <div class="preview-panel">
          <div class="preview-header">
            <h3>{{ activeSlide ? `Slide #${activeSlide.index + 1} 对照预览` : 'Slide 预览' }}</h3>
            <div class="slide-nav" aria-label="Slide navigation">
              <button type="button" :disabled="!canGoPrevious" @click="goToFirstSlide">首页</button>
              <button type="button" :disabled="!canGoPrevious" @click="goToPreviousSlide">上一页</button>
              <span>{{ activeSlide ? activeSlide.index + 1 : 0 }} / {{ slideCount }}</span>
              <button type="button" :disabled="!canGoNext" @click="goToNextSlide">下一页</button>
              <button type="button" :disabled="!canGoNext" @click="goToLastSlide">末页</button>
            </div>
          </div>

          <div class="preview-grid">
            <section class="preview-column" aria-labelledby="svg-preview-title">
              <div class="preview-label-row">
                <p id="svg-preview-title" class="preview-label">SVG Renderer</p>
                <span class="preview-meta">字符串拼装 · DOM 挂载</span>
              </div>
              <div v-if="activeSlideSvg" class="slide-preview-host" v-html="activeSlideSvg" />
              <p v-else class="empty-text">暂无可预览的 slide。</p>
            </section>

            <section class="preview-column" aria-labelledby="canvas-preview-title">
              <div class="preview-label-row">
                <p id="canvas-preview-title" class="preview-label">Canvas Renderer</p>
                <span class="preview-meta">2D context · 即时绘制</span>
              </div>
              <div class="slide-preview-host canvas-preview-host">
                <canvas ref="canvasHost" class="slide-preview-canvas" />
              </div>
            </section>
          </div>
        </div>

        <div class="slide-list">
          <h3>Slides</h3>
          <ol>
            <li v-for="slide in presentation.slides" :key="slide.relationshipId">
              <button
                type="button"
                class="slide-heading"
                :class="{ 'slide-heading-active': activeSlide?.index === slide.index }"
                :aria-current="activeSlide?.index === slide.index ? 'true' : undefined"
                @click="goToSlide(slide.index)"
              >
                <span>#{{ slide.index + 1 }}</span>
                <code>{{ slide.part }}</code>
              </button>

              <div class="element-stats" aria-label="Slide element summary">
                <span>文本 {{ slide.elements.filter((element) => element.type === 'text').length }}</span>
                <span>图片 {{ slide.elements.filter((element) => element.type === 'image').length }}</span>
                <span>形状 {{ slide.elements.filter((element) => element.type === 'shape').length }}</span>
                <span>连接符 {{ slide.elements.filter((element) => element.type === 'connector').length }}</span>
                <span>未知 {{ slide.elements.filter((element) => element.type === 'unknown').length }}</span>
              </div>

              <ul v-if="slide.elements.length > 0" class="element-list">
                <li v-for="element in slide.elements" :key="element.id">
                  <strong>{{ element.type }}</strong>
                  <span v-if="element.type === 'text'">{{ element.text }}</span>
                  <code v-else-if="element.type === 'image'">{{ element.imagePart ?? element.relationshipId ?? 'missing image relationship' }}</code>
                  <code v-else-if="element.type === 'shape'">
                    fill {{ fillLabel(element.fill) }} · line
                    {{ element.line?.color ?? 'none' }} {{ element.line?.width ?? 0 }}px
                  </code>
                  <code v-else-if="element.type === 'connector'">
                    line {{ element.line?.color ?? solidFillColor(element.fill, 'none') }}
                    {{ element.line?.width ?? 0 }}px
                  </code>
                  <span v-else>{{ element.nodeName }}</span>
                </li>
              </ul>
            </li>
          </ol>
        </div>
      </div>

      <div v-else class="status-card">请选择一个样本或上传 PPTX 文件开始解析。</div>
    </section>

    <section class="diagnostics-panel">
      <div class="diagnostics-header">
        <div>
          <h2>Diagnostics</h2>
          <p class="diagnostics-summary">
            共 {{ diagnostics.length }} 条 · info {{ diagnosticCounts.info }} · warning {{ diagnosticCounts.warning }} · error
            {{ diagnosticCounts.error }}
          </p>
        </div>

        <div v-if="diagnostics.length > 0" class="diagnostics-filters" aria-label="Diagnostics filters">
          <label>
            Severity
            <select v-model="diagnosticSeverityFilter">
              <option value="all">全部</option>
              <option value="info">info</option>
              <option value="warning">warning</option>
              <option value="error">error</option>
            </select>
          </label>

          <label>
            Code
            <select v-model="diagnosticCodeFilter">
              <option value="all">全部</option>
              <option v-for="code in diagnosticCodes" :key="code" :value="code">{{ code }}</option>
            </select>
          </label>

          <label>
            Slide
            <select v-model="diagnosticSlideFilter">
              <option value="all">全部</option>
              <option v-for="slideIndex in diagnosticSlideOptions" :key="slideIndex" :value="String(slideIndex)">
                #{{ slideIndex + 1 }}
              </option>
            </select>
          </label>
        </div>
      </div>

      <p v-if="diagnostics.length === 0" class="empty-text">暂无 diagnostics。</p>
      <p v-else-if="filteredDiagnostics.length === 0" class="empty-text">当前过滤条件下暂无 diagnostics。</p>
      <ul v-else>
        <li v-for="(diagnostic, index) in filteredDiagnostics" :key="`${diagnostic.code}-${diagnostic.part ?? 'global'}-${index}`">
          <div class="diagnostic-title-row">
            <strong :class="`severity-${diagnostic.severity}`">{{ diagnostic.severity }}</strong>
            <code>{{ diagnostic.code }}</code>
            <span>{{ diagnostic.message }}</span>
          </div>

          <dl class="diagnostic-context">
            <div v-if="diagnostic.slideIndex !== undefined">
              <dt>Slide</dt>
              <dd>#{{ diagnostic.slideIndex + 1 }}</dd>
            </div>
            <div v-if="diagnostic.elementId">
              <dt>Element</dt>
              <dd>{{ diagnostic.elementId }}</dd>
            </div>
            <div v-if="diagnostic.part">
              <dt>Part</dt>
              <dd>{{ diagnostic.part }}</dd>
            </div>
          </dl>

          <pre v-if="formatDiagnosticDetail(diagnostic.detail)" class="diagnostic-detail">{{ formatDiagnosticDetail(diagnostic.detail) }}</pre>
        </li>
      </ul>
    </section>
  </main>
</template>

<style scoped>
:global(*) {
  box-sizing: border-box;
}

:global(body) {
  margin: 0;
  min-width: 320px;
  background: #0f172a;
  color: #e2e8f0;
  font-family:
    Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
}

button,
input {
  font: inherit;
}

.app-shell {
  display: grid;
  gap: 24px;
  width: min(1240px, calc(100vw - 32px));
  margin: 0 auto;
  padding: 48px 0;
}

.hero-panel,
.control-panel,
.result-panel,
.diagnostics-panel {
  border: 1px solid rgba(148, 163, 184, 0.24);
  border-radius: 24px;
  background: rgba(15, 23, 42, 0.82);
  box-shadow: 0 24px 80px rgba(15, 23, 42, 0.36);
}

.hero-panel {
  padding: 40px;
  background:
    radial-gradient(circle at top left, rgba(56, 189, 248, 0.22), transparent 36%),
    linear-gradient(135deg, rgba(30, 41, 59, 0.92), rgba(15, 23, 42, 0.92));
}

.eyebrow {
  margin: 0 0 10px;
  color: #38bdf8;
  font-size: 13px;
  font-weight: 700;
  letter-spacing: 0.12em;
  text-transform: uppercase;
}

h1,
h2,
h3,
p {
  margin-top: 0;
}

h1 {
  max-width: 760px;
  margin-bottom: 16px;
  font-size: clamp(36px, 6vw, 64px);
  line-height: 1;
}

.description {
  max-width: 760px;
  margin-bottom: 0;
  color: #cbd5e1;
  font-size: 18px;
  line-height: 1.7;
}

.control-panel,
.result-panel,
.diagnostics-panel {
  padding: 28px;
}

.actions {
  display: grid;
  grid-template-columns: minmax(240px, 320px) 1fr;
  gap: 20px;
}

.upload-card {
  display: grid;
  gap: 12px;
  align-content: center;
  min-height: 132px;
  padding: 20px;
  border: 1px dashed #38bdf8;
  border-radius: 18px;
  background: rgba(14, 165, 233, 0.1);
  color: #e0f2fe;
  cursor: pointer;
}

.upload-card input {
  max-width: 100%;
}

.sample-list {
  display: flex;
  flex-wrap: wrap;
  gap: 12px;
  align-content: flex-start;
}

.sample-list button {
  border: 1px solid rgba(148, 163, 184, 0.32);
  border-radius: 999px;
  padding: 10px 16px;
  background: rgba(30, 41, 59, 0.9);
  color: #e2e8f0;
  cursor: pointer;
}

.sample-list button:disabled,
.upload-card:has(input:disabled) {
  cursor: wait;
  opacity: 0.6;
}

.status-card,
.presentation-card {
  border-radius: 18px;
  padding: 20px;
  background: rgba(30, 41, 59, 0.72);
}

.error-card {
  border: 1px solid rgba(248, 113, 113, 0.48);
  color: #fecaca;
}

.metrics-grid {
  display: grid;
  grid-template-columns: repeat(4, minmax(0, 1fr));
  gap: 14px;
  margin: 24px 0;
}

.metrics-grid div {
  min-width: 0;
  border-radius: 16px;
  padding: 16px;
  background: rgba(15, 23, 42, 0.86);
}

dt {
  color: #94a3b8;
  font-size: 13px;
}

dd {
  overflow-wrap: anywhere;
  margin: 8px 0 0;
  font-size: 20px;
  font-weight: 800;
}

.preview-panel {
  display: grid;
  gap: 14px;
  margin-bottom: 24px;
}

.preview-header,
.slide-nav,
.diagnostics-header,
.diagnostics-filters,
.diagnostic-title-row {
  display: flex;
  flex-wrap: wrap;
  gap: 12px;
  align-items: center;
  justify-content: space-between;
}

.preview-header h3 {
  margin-bottom: 0;
}

.slide-nav {
  justify-content: flex-start;
}

.slide-nav button {
  border: 1px solid rgba(56, 189, 248, 0.42);
  border-radius: 999px;
  padding: 8px 14px;
  background: rgba(14, 165, 233, 0.14);
  color: #e0f2fe;
  cursor: pointer;
}

.slide-nav button:disabled {
  cursor: not-allowed;
  opacity: 0.45;
}

.slide-nav span {
  color: #cbd5e1;
  font-weight: 700;
}

.diagnostics-header {
  margin-bottom: 18px;
}

.diagnostics-header h2 {
  margin-bottom: 8px;
}

.diagnostics-summary {
  margin-bottom: 0;
  color: #cbd5e1;
  font-size: 14px;
}

.diagnostics-filters {
  justify-content: flex-start;
}

.diagnostics-filters label {
  display: grid;
  gap: 6px;
  color: #94a3b8;
  font-size: 12px;
  font-weight: 700;
  letter-spacing: 0.04em;
  text-transform: uppercase;
}

.diagnostics-filters select {
  min-width: 132px;
  border: 1px solid rgba(148, 163, 184, 0.32);
  border-radius: 10px;
  padding: 8px 10px;
  background: rgba(15, 23, 42, 0.92);
  color: #e2e8f0;
}

.diagnostic-title-row {
  justify-content: flex-start;
}

.diagnostic-context {
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
  margin: 0;
}

.diagnostic-context div {
  display: flex;
  gap: 6px;
  border-radius: 999px;
  padding: 4px 10px;
  background: rgba(148, 163, 184, 0.12);
}

.diagnostic-context dt,
.diagnostic-context dd {
  margin: 0;
  font-size: 12px;
}

.diagnostic-context dt {
  color: #94a3b8;
}

.diagnostic-context dd {
  color: #e2e8f0;
  font-weight: 700;
}

.diagnostic-detail {
  overflow: auto;
  max-width: 100%;
  margin: 0;
  border-radius: 10px;
  padding: 10px;
  background: rgba(2, 6, 23, 0.72);
  color: #c4b5fd;
  font-size: 12px;
}

.preview-grid {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 16px;
}

.preview-column {
  display: grid;
  gap: 10px;
}

.preview-label-row {
  display: flex;
  flex-wrap: wrap;
  gap: 8px 12px;
  align-items: baseline;
  justify-content: space-between;
}

.preview-label {
  margin-bottom: 0;
  color: #e2e8f0;
  font-size: 14px;
  font-weight: 700;
  letter-spacing: 0.04em;
  text-transform: uppercase;
}

.preview-meta {
  color: #94a3b8;
  font-size: 12px;
}

.slide-preview-host {
  display: grid;
  place-items: center;
  min-height: 220px;
  border-radius: 16px;
  overflow: hidden;
  padding: 14px;
  background:
    linear-gradient(180deg, rgba(226, 232, 240, 0.9), rgba(203, 213, 225, 0.95));
}

.canvas-preview-host {
  align-items: start;
}

:deep(.slide-preview) {
  display: block;
  width: 100%;
  max-height: 520px;
}

.slide-preview-canvas {
  display: block;
  width: 100%;
  height: auto;
  max-height: 520px;
  background: #f8fafc;
}

.slide-list ol,
.diagnostics-panel ul {
  display: grid;
  gap: 10px;
  padding: 0;
  list-style: none;
}

.slide-list > ol > li,
.diagnostics-panel li {
  display: grid;
  gap: 10px;
  border-radius: 12px;
  padding: 12px;
  background: rgba(15, 23, 42, 0.72);
}

.slide-heading,
.element-stats,
.element-list li {
  display: flex;
  flex-wrap: wrap;
  gap: 10px;
  align-items: center;
}

.slide-heading {
  width: 100%;
  border: 1px solid transparent;
  border-radius: 10px;
  padding: 8px 10px;
  background: transparent;
  color: inherit;
  text-align: left;
  cursor: pointer;
}

.slide-heading-active {
  border-color: rgba(56, 189, 248, 0.56);
  background: rgba(56, 189, 248, 0.12);
}

.element-stats span {
  border-radius: 999px;
  padding: 4px 10px;
  background: rgba(56, 189, 248, 0.12);
  color: #bae6fd;
  font-size: 13px;
}

.element-list {
  display: grid;
  gap: 8px;
  margin: 0;
  padding: 0;
  list-style: none;
}

.element-list li {
  border-radius: 10px;
  padding: 10px;
  background: rgba(30, 41, 59, 0.68);
}

code {
  color: #bae6fd;
  overflow-wrap: anywhere;
}

small {
  color: #94a3b8;
}

.empty-text {
  margin-bottom: 0;
  color: #94a3b8;
}

.severity-info {
  color: #93c5fd;
}

.severity-warning {
  color: #facc15;
}

.severity-error {
  color: #fca5a5;
}

@media (max-width: 960px) {
  .preview-grid,
  .actions,
  .metrics-grid {
    grid-template-columns: 1fr;
  }
}

@media (max-width: 780px) {
  .hero-panel,
  .control-panel,
  .result-panel,
  .diagnostics-panel {
    padding: 22px;
  }
}
</style>
