#!/usr/bin/env node
import { chromium } from '@playwright/test'
import { PNG } from 'pngjs'
import pixelmatch from 'pixelmatch'
import { mkdir, readFile, writeFile } from 'node:fs/promises'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const repositoryRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const defaults = {
  baseUrl: 'http://localhost:5175/',
  sample: 'Fixture 4b00',
  slide: 1,
  renderer: 'svg',
  reference: '',
  referenceCrop: '',
  outputDir: 'reports/visual-regression',
  pixelThreshold: 0.1,
  maxMismatchRatio: 0.05,
}

const options = parseArgs(process.argv.slice(2))
const cases = await loadCases(options)
const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1440, height: 1100 }, deviceScaleFactor: 1 })
const reports = []

try {
  await page.goto(options.baseUrl ?? defaults.baseUrl, { waitUntil: 'networkidle' })

  for (const testCase of cases) {
    reports.push(await runCase(page, testCase))
  }

  const outputDir = resolve(repositoryRoot, options.outputDir ?? defaults.outputDir)
  await mkdir(outputDir, { recursive: true })
  const summaryPath = resolve(outputDir, 'summary.json')
  const summary = {
    passed: reports.every((report) => report.passed),
    total: reports.length,
    failed: reports.filter((report) => !report.passed).length,
    reports,
  }

  await writeFile(summaryPath, `${JSON.stringify(summary, null, 2)}\n`)

  if (reports.length === 1) {
    console.log(JSON.stringify({ ...reports[0], summary: relativePath(summaryPath) }, null, 2))
  } else {
    console.log(JSON.stringify({ ...summary, summary: relativePath(summaryPath) }, null, 2))
  }

  if (!summary.passed) {
    process.exitCode = 1
  }
} finally {
  await browser.close()
}

async function loadCases(options) {
  if (options.cases) {
    const casesPath = resolve(repositoryRoot, String(options.cases))
    const content = await readFile(casesPath, 'utf8')
    const parsed = JSON.parse(content)
    const caseList = Array.isArray(parsed) ? parsed : parsed.cases

    if (!Array.isArray(caseList) || caseList.length === 0) {
      printUsageAndExit('--cases 文件必须包含非空数组或 { "cases": [...] }。')
    }

    return caseList.flatMap((testCase) => normalizeCase({ ...defaults, ...options, ...testCase }))
  }

  const config = { ...defaults, ...options }

  if (!config.reference) {
    printUsageAndExit('缺少 WPS 参考图，请传入 --reference <path> 或 --cases <json>。')
  }

  return normalizeCase(config)
}

function normalizeCase(config) {
  const renderers = config.renderers ?? config.renderer
  const rendererList = Array.isArray(renderers) ? renderers : String(renderers).split(',').map((renderer) => renderer.trim())

  return rendererList.map((renderer) => {
    if (!['svg', 'canvas'].includes(renderer)) {
      printUsageAndExit('--renderer/renderers 只能是 svg 或 canvas。')
    }

    return {
      ...config,
      renderer,
      slide: Number(config.slide),
      pixelThreshold: Number(config.pixelThreshold),
      maxMismatchRatio: Number(config.maxMismatchRatio),
      referenceCrop: parseCrop(config.referenceCrop),
    }
  })
}

async function runCase(page, config) {
  await page.goto(config.baseUrl, { waitUntil: 'networkidle' })
  await page.getByRole('button', { name: config.sample, exact: true }).click()
  await page.waitForSelector('.presentation-card', { timeout: 30000 })
  await selectSlide(page, Number(config.slide))
  await page.waitForTimeout(300)

  const selector = config.renderer === 'svg' ? '[aria-labelledby="svg-preview-title"] svg.slide-preview' : '[aria-labelledby="canvas-preview-title"] canvas'
  const actualBuffer = await captureSlide(page, selector)
  const outputDir = resolve(repositoryRoot, config.outputDir)
  await mkdir(outputDir, { recursive: true })

  const prefix = `${slugify(config.sample)}-slide-${config.slide}-${config.renderer}`
  const actualPath = resolve(outputDir, `${prefix}-actual.png`)
  const diffPath = resolve(outputDir, `${prefix}-diff.png`)
  const reportPath = resolve(outputDir, `${prefix}-report.json`)

  await writeFile(actualPath, actualBuffer)

  const result = await comparePngs(resolve(repositoryRoot, config.reference), actualBuffer, diffPath, {
    pixelThreshold: config.pixelThreshold,
    maxMismatchRatio: config.maxMismatchRatio,
    referenceCrop: config.referenceCrop,
  })
  const report = {
    sample: config.sample,
    slide: Number(config.slide),
    renderer: config.renderer,
    reference: config.reference,
    actual: relativePath(actualPath),
    diff: relativePath(diffPath),
    report: relativePath(reportPath),
    ...result,
  }

  await writeFile(reportPath, `${JSON.stringify(report, null, 2)}\n`)
  return report
}

async function selectSlide(page, slideNumber) {
  const currentText = await page.locator('.slide-nav span').textContent()
  const total = Number(currentText?.split('/')[1]?.trim() ?? 0)

  if (!Number.isInteger(slideNumber) || slideNumber < 1 || slideNumber > total) {
    throw new Error(`slide 超出范围：${slideNumber}，当前样本共有 ${total} 页。`)
  }

  for (let current = 1; current < slideNumber; current += 1) {
    await page.getByRole('button', { name: '下一页' }).click()
  }

  await page.waitForFunction((expected) => document.querySelector('.slide-nav span')?.textContent?.trim().startsWith(`${expected} /`), slideNumber)
}

async function captureSlide(page, selector) {
  const locator = page.locator(selector).first()
  await locator.waitFor({ timeout: 30000 })

  const actualBuffer = await locator.screenshot()
  const actual = PNG.sync.read(actualBuffer)
  const intrinsicSize = await locator.evaluate((element) => {
    if (element instanceof HTMLCanvasElement) {
      return { width: element.width, height: element.height }
    }

    if (element instanceof SVGSVGElement) {
      return {
        width: Number(element.getAttribute('width')) || Math.ceil(element.viewBox.baseVal.width),
        height: Number(element.getAttribute('height')) || Math.ceil(element.viewBox.baseVal.height),
      }
    }

    const box = element.getBoundingClientRect()
    return { width: Math.ceil(box.width), height: Math.ceil(box.height) }
  })

  if (actual.width === intrinsicSize.width && actual.height === intrinsicSize.height) {
    return actualBuffer
  }

  return PNG.sync.write(resizePng(actual, intrinsicSize.width, intrinsicSize.height))
}

async function comparePngs(referencePath, actualBuffer, diffPath, options) {
  const sourceReference = PNG.sync.read(await readFile(referencePath))
  const reference = cropPng(sourceReference, options.referenceCrop)
  const actual = PNG.sync.read(actualBuffer)
  const normalizedReference = resizePng(reference, actual.width, actual.height)
  const diff = new PNG({ width: actual.width, height: actual.height })
  const mismatchedPixels = pixelmatch(normalizedReference.data, actual.data, diff.data, actual.width, actual.height, { threshold: options.pixelThreshold })
  const totalPixels = actual.width * actual.height
  const mismatchRatio = totalPixels === 0 ? 0 : mismatchedPixels / totalPixels

  await writeFile(diffPath, PNG.sync.write(diff))

  return {
    passed: mismatchRatio <= options.maxMismatchRatio,
    pixelThreshold: options.pixelThreshold,
    maxMismatchRatio: options.maxMismatchRatio,
    mismatchedPixels,
    totalPixels,
    mismatchRatio,
    comparedWidth: actual.width,
    comparedHeight: actual.height,
    referenceWidth: sourceReference.width,
    referenceHeight: sourceReference.height,
    referenceCropWidth: reference.width,
    referenceCropHeight: reference.height,
    actualWidth: actual.width,
    actualHeight: actual.height,
  }
}

function cropPng(source, crop) {
  if (!crop) {
    return source
  }

  const x = Math.max(0, Math.min(source.width - 1, crop.x))
  const y = Math.max(0, Math.min(source.height - 1, crop.y))
  const width = Math.max(1, Math.min(source.width - x, crop.width))
  const height = Math.max(1, Math.min(source.height - y, crop.height))
  const target = new PNG({ width, height })

  for (let targetY = 0; targetY < height; targetY += 1) {
    for (let targetX = 0; targetX < width; targetX += 1) {
      const sourceIndex = ((y + targetY) * source.width + x + targetX) * 4
      const targetIndex = (targetY * width + targetX) * 4

      target.data[targetIndex] = source.data[sourceIndex]
      target.data[targetIndex + 1] = source.data[sourceIndex + 1]
      target.data[targetIndex + 2] = source.data[sourceIndex + 2]
      target.data[targetIndex + 3] = source.data[sourceIndex + 3]
    }
  }

  return target
}

function resizePng(source, width, height) {
  if (source.width === width && source.height === height) {
    return source
  }

  const target = new PNG({ width, height })

  for (let y = 0; y < height; y += 1) {
    const sourceY = Math.min(source.height - 1, Math.floor((y / height) * source.height))

    for (let x = 0; x < width; x += 1) {
      const sourceX = Math.min(source.width - 1, Math.floor((x / width) * source.width))
      const sourceIndex = (sourceY * source.width + sourceX) * 4
      const targetIndex = (y * width + x) * 4

      target.data[targetIndex] = source.data[sourceIndex]
      target.data[targetIndex + 1] = source.data[sourceIndex + 1]
      target.data[targetIndex + 2] = source.data[sourceIndex + 2]
      target.data[targetIndex + 3] = source.data[sourceIndex + 3]
    }
  }

  return target
}

function parseArgs(args) {
  const parsed = {}

  for (let index = 0; index < args.length; index += 1) {
    const argument = args[index]

    if (!argument.startsWith('--')) {
      continue
    }

    const key = argument.slice(2)
    const value = args[index + 1]

    if (!value || value.startsWith('--')) {
      parsed[key] = true
      continue
    }

    parsed[key] = value
    index += 1
  }

  return parsed
}

function printUsageAndExit(message) {
  console.error(message)
  console.error('用法：pnpm visual:compare -- --reference <wps.png> [--referenceCrop x,y,width,height] [--baseUrl http://localhost:5175/] [--sample "Fixture 4b00"] [--slide 8] [--renderer svg|canvas] [--pixelThreshold 0.1] [--maxMismatchRatio 0.05]')
  console.error('批量：pnpm visual:compare -- --cases scripts/visual-regression-cases.json')
  process.exit(1)
}

function parseCrop(value) {
  if (!value) {
    return undefined
  }

  if (typeof value === 'object') {
    return value
  }

  const values = String(value).split(',').map((item) => Number(item.trim()))

  if (values.length !== 4 || values.some((item) => !Number.isFinite(item))) {
    printUsageAndExit('--referenceCrop 格式必须是 x,y,width,height。')
  }

  const [x, y, width, height] = values.map((item) => Math.round(item))

  if (width <= 0 || height <= 0) {
    printUsageAndExit('--referenceCrop 的 width/height 必须大于 0。')
  }

  return { x, y, width, height }
}

function slugify(value) {
  return value.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '')
}

function relativePath(path) {
  return path.startsWith(repositoryRoot) ? path.slice(repositoryRoot.length + 1) : path
}
