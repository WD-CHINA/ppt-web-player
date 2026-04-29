<script setup lang="ts">
import { computed, ref } from 'vue'
import { parsePptx, type Diagnostic, type Presentation } from '@pptx-player/core'

import { sampleDecks } from './samples'

const isLoading = ref(false)
const errorMessage = ref('')
const diagnostics = ref<Diagnostic[]>([])
const presentation = ref<Presentation | null>(null)
const activeFileName = ref('')

const hasResult = computed(() => presentation.value !== null)
const firstSlide = computed(() => presentation.value?.slides[0] ?? null)
const previewElements = computed(() => firstSlide.value?.elements.filter((element) => element.transform) ?? [])

async function parseInput(input: Blob | ArrayBuffer, fileName: string) {
  isLoading.value = true
  errorMessage.value = ''
  presentation.value = null
  diagnostics.value = []
  activeFileName.value = fileName

  try {
    const result = await parsePptx(input)
    presentation.value = result.presentation
    diagnostics.value = result.diagnostics

    if (!result.presentation) {
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
</script>

<template>
  <main class="app-shell">
    <section class="hero-panel">
      <p class="eyebrow">PPTX Web Player · Phase 0</p>
      <h1>企业可控 PPTX Web Player</h1>
      <p class="description">
        当前阶段打通 PPTX 解包、XML 解析、relationships 解析、presentation 尺寸和 slide 列表读取。
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
          <h3>第一页预览</h3>
          <svg
            v-if="firstSlide"
            class="slide-preview"
            role="img"
            :aria-label="`${firstSlide.part} preview`"
            :viewBox="`0 0 ${presentation.width} ${presentation.height}`"
          >
            <rect width="100%" height="100%" rx="12" :fill="firstSlide.background?.color ?? '#f8fafc'" />
            <g v-for="element in previewElements" :key="element.id">
              <template v-if="element.transform">
                <text
                  v-if="element.type === 'text'"
                  :x="element.transform.x"
                  :y="element.transform.y + 16"
                  :font-size="Math.max(10, Math.min(24, element.transform.height / 2))"
                  fill="#0f172a"
                >
                  {{ element.text.slice(0, 48) }}
                </text>
                <g v-else-if="element.type === 'image'">
                  <rect
                    :x="element.transform.x"
                    :y="element.transform.y"
                    :width="element.transform.width"
                    :height="element.transform.height"
                    rx="8"
                    fill="#38bdf8"
                    fill-opacity="0.28"
                    stroke="#0284c7"
                  />
                  <text :x="element.transform.x + 8" :y="element.transform.y + 20" font-size="12" fill="#075985">
                    {{ element.part ?? 'image' }}
                  </text>
                </g>
                <rect
                  v-else-if="element.type === 'shape'"
                  :x="element.transform.x"
                  :y="element.transform.y"
                  :width="element.transform.width"
                  :height="element.transform.height"
                  rx="6"
                  :fill="element.fill?.color ?? '#cbd5e1'"
                  fill-opacity="0.7"
                  :stroke="element.line?.color ?? '#334155'"
                  :stroke-width="element.line?.width ?? 1"
                  stroke-opacity="0.45"
                />
                <line
                  v-else-if="element.type === 'connector'"
                  :x1="element.transform.x"
                  :y1="element.transform.y + element.transform.height / 2"
                  :x2="element.transform.x + element.transform.width"
                  :y2="element.transform.y + element.transform.height / 2"
                  :stroke="element.line?.color ?? element.fill?.color ?? '#64748b'"
                  :stroke-width="element.line?.width ?? 4"
                  stroke-linecap="round"
                />
                <rect
                  v-else
                  :x="element.transform.x"
                  :y="element.transform.y"
                  :width="element.transform.width"
                  :height="element.transform.height"
                  fill="none"
                  stroke="#f97316"
                  stroke-dasharray="8 6"
                />
              </template>
            </g>
          </svg>
          <p v-else class="empty-text">暂无可预览的 slide。</p>
        </div>

        <div class="slide-list">
          <h3>Slides</h3>
          <ol>
            <li v-for="slide in presentation.slides" :key="slide.relationshipId">
              <div class="slide-heading">
                <span>#{{ slide.index + 1 }}</span>
                <code>{{ slide.part }}</code>
              </div>

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
                  <code v-else-if="element.type === 'image'">{{ element.part ?? element.relationshipId ?? 'missing image relationship' }}</code>
                  <code v-else-if="element.type === 'shape'">
                    fill {{ element.fill?.color ?? 'none' }} · line
                    {{ element.line?.color ?? 'none' }} {{ element.line?.width ?? 0 }}px
                  </code>
                  <code v-else-if="element.type === 'connector'">
                    line {{ element.line?.color ?? element.fill?.color ?? 'none' }}
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
      <h2>Diagnostics</h2>
      <p v-if="diagnostics.length === 0" class="empty-text">暂无 diagnostics。</p>
      <ul v-else>
        <li v-for="(diagnostic, index) in diagnostics" :key="`${diagnostic.code}-${index}`">
          <strong :class="`severity-${diagnostic.severity}`">{{ diagnostic.severity }}</strong>
          <code>{{ diagnostic.code }}</code>
          <span>{{ diagnostic.message }}</span>
          <small v-if="diagnostic.part">{{ diagnostic.part }}</small>
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
  width: min(1120px, calc(100vw - 32px));
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
  max-width: 720px;
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

.slide-preview {
  width: 100%;
  max-height: 520px;
  border-radius: 16px;
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

@media (max-width: 780px) {
  .actions,
  .metrics-grid {
    grid-template-columns: 1fr;
  }

  .hero-panel,
  .control-panel,
  .result-panel,
  .diagnostics-panel {
    padding: 22px;
  }
}
</style>
