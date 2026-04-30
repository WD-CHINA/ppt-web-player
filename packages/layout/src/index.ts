import type { Paragraph, TextElement, TextSpacing, TextStyle } from '@pptx-player/core'

export const layoutPackageName = '@pptx-player/layout'

export type TextHorizontalAlign = 'left' | 'center' | 'right'

export interface TextLayoutRun {
  text: string
  style?: TextStyle
  fontSize?: number
}

export interface TextLayoutLine {
  x: number
  y: number
  runs: TextLayoutRun[]
}

export interface TextLayoutBlock {
  defaultFontSize: number
  align: TextHorizontalAlign
  lines: TextLayoutLine[]
}

export interface TextMeasureContext {
  text: string
  style: TextStyle | undefined
  defaultFontSize: number
}

export interface LayoutTextElementOptions {
  measureText?: (context: TextMeasureContext) => number
}

interface BulletLayoutContext {
  text: string
  style: TextStyle | undefined
  width: number
}

interface AutoNumberState {
  nextByLevel: Map<number, number>
}

interface TextContentBox {
  x: number
  y: number
  width: number
  height: number
}

interface ParagraphLayoutResult {
  block: TextLayoutBlock
  height: number
}

export function layoutTextElement(element: TextElement, options: LayoutTextElementOptions = {}): TextLayoutBlock | null {
  const transform = element.transform

  if (!transform) {
    return null
  }

  const paragraphs = normalizedParagraphs(element)
  const baseDefaultFontSize = baseFontSize(element)
  const contentBox = textContentBox(element, baseDefaultFontSize)
  const align = paragraphAlign(paragraphs)
  const layout = layoutParagraphs(element, paragraphs, contentBox, align, baseDefaultFontSize, 1, options)
  const autoFit = element.textBody.properties?.autoFit

  if (autoFit?.type !== 'normal' || layout.height <= contentBox.height || contentBox.height <= 0) {
    return applyVerticalAnchor(layout.block, layout.height, contentBox, element.textBody.properties?.verticalAnchor)
  }

  const minScale = autoFit.fontScale !== undefined ? Math.max(0.1, Math.min(1, autoFit.fontScale / 100000)) : 0.6
  const scale = Math.max(minScale, Math.min(1, contentBox.height / layout.height))
  const fittedLayout = layoutParagraphs(element, paragraphs, contentBox, align, baseDefaultFontSize, scale, options)

  return applyVerticalAnchor(fittedLayout.block, fittedLayout.height, contentBox, element.textBody.properties?.verticalAnchor)
}

function layoutParagraphs(
  element: TextElement,
  paragraphs: Paragraph[],
  contentBox: TextContentBox,
  align: TextHorizontalAlign,
  defaultFontSize: number,
  fontScale: number,
  options: LayoutTextElementOptions,
): ParagraphLayoutResult {
  const scaledDefaultFontSize = defaultFontSize * fontScale
  const blockX = contentBox.x
  const blockY = contentBox.y + scaledDefaultFontSize
  const blockWidth = Math.max(contentBox.width, scaledDefaultFontSize)
  const canWrap = element.textBody.properties?.wrap !== false
  let yOffset = 0
  const lines: TextLayoutLine[] = []
  const autoNumberState: AutoNumberState = {
    nextByLevel: new Map(),
  }

  for (const paragraph of paragraphs) {
    const bullet = paragraphBullet(paragraph, autoNumberState)
    const bulletRunStyle = scaledTextStyle(bulletStyle(paragraph), fontScale)
    const bulletLayout = measureBulletLayout(bullet, bulletRunStyle, scaledDefaultFontSize, options.measureText)
    const paragraphLines = splitParagraphLines(paragraph, fontScale, scaledDefaultFontSize)
    const wrappedLines = paragraphLines.flatMap((lineRuns, lineIndex) => {
      if (!canWrap) {
        return [lineRuns]
      }

      return wrapLineRuns(lineRuns, {
        availableWidth: paragraphAvailableWidth(paragraph, blockWidth, lineIndex, bulletLayout),
        defaultFontSize: scaledDefaultFontSize,
        measureText: options.measureText,
      })
    })
    const lineHeight = paragraphLineHeight(paragraph, scaledDefaultFontSize)
    const spaceBefore = paragraphSpaceBefore(paragraph, scaledDefaultFontSize)

    wrappedLines.forEach((lineRuns, lineIndex) => {
      const contentRuns = lineRuns.length > 0 ? lineRuns : [{ text: '', style: scaledTextStyle(paragraph.style?.defaultRunStyle, fontScale), fontSize: scaledDefaultFontSize }]
      const runs = bulletLayout && lineIndex === 0 ? [{ text: bulletLayout.text, style: bulletLayout.style, fontSize: scaledDefaultFontSize }, ...contentRuns] : contentRuns
      const x = paragraphLineX(paragraph, lineIndex, blockX, blockWidth, align, runs, scaledDefaultFontSize, options.measureText, bulletLayout)
      const y = blockY + yOffset + spaceBefore + lineIndex * lineHeight
      lines.push({ x, y, runs })
    })

    yOffset += paragraphBlockHeight(paragraph, wrappedLines, scaledDefaultFontSize)
  }

  return {
    block: {
      defaultFontSize: scaledDefaultFontSize,
      align,
      lines,
    },
    height: yOffset,
  }
}

export function textStyleFontSize(style: TextStyle | undefined, defaultFontSize: number): number {
  return style?.fontSize ? normalizedFontSize(style.fontSize) : defaultFontSize
}

export function normalizeTextRunText(text: string): string {
  return text.split('\t').join('    ')
}

export function normalizedFontSize(fontSize: number): number {
  return Math.max(1, fontSize / 100)
}

function applyVerticalAnchor(
  block: TextLayoutBlock,
  height: number,
  contentBox: TextContentBox,
  verticalAnchor: 'top' | 'middle' | 'bottom' | undefined,
): TextLayoutBlock {
  if (!verticalAnchor || verticalAnchor === 'top' || height >= contentBox.height || contentBox.height <= 0) {
    return block
  }

  const offset = verticalAnchor === 'middle' ? (contentBox.height - height) / 2 : contentBox.height - height

  return {
    ...block,
    lines: block.lines.map((line) => ({ ...line, y: line.y + offset })),
  }
}

function normalizedParagraphs(element: TextElement): Paragraph[] {
  return element.textBody.paragraphs.length > 0 ? element.textBody.paragraphs : [{ runs: [{ text: element.text }], text: element.text }]
}

function textContentBox(element: TextElement, defaultFontSize: number): TextContentBox {
  const transform = element.transform
  const inset = element.textBody.properties?.inset
  const left = inset?.left ?? 0
  const top = inset?.top ?? 0
  const right = inset?.right ?? 0
  const bottom = inset?.bottom ?? 0

  return {
    x: transform?.x ? transform.x + left : left,
    y: transform?.y ? transform.y + top : top,
    width: Math.max(defaultFontSize, (transform?.width ?? defaultFontSize) - left - right),
    height: Math.max(0, (transform?.height ?? defaultFontSize) - top - bottom),
  }
}

function scaledTextStyle(style: TextStyle | undefined, fontScale: number): TextStyle | undefined {
  if (!style || fontScale === 1 || style.fontSize === undefined) {
    return style
  }

  return {
    ...style,
    fontSize: style.fontSize * fontScale,
  }
}

function splitParagraphLines(paragraph: Paragraph, fontScale: number, defaultFontSize: number): TextLayoutRun[][] {
  const lines: TextLayoutRun[][] = [[]]

  for (const run of paragraph.runs) {
    const segments = run.text.split('\n')
    const style = scaledTextStyle(run.style, fontScale)

    segments.forEach((segment, index) => {
      const currentLine = lines[lines.length - 1]

      if (segment.length > 0 && currentLine) {
        currentLine.push({ text: segment, style, fontSize: textStyleFontSize(style, defaultFontSize) })
      }

      if (index < segments.length - 1) {
        lines.push([])
      }
    })
  }

  return lines
}

function wrapLineRuns(
  lineRuns: TextLayoutRun[],
  options: {
    availableWidth: number
    defaultFontSize: number
    measureText?: (context: TextMeasureContext) => number
  },
): TextLayoutRun[][] {
  if (lineRuns.length === 0 || options.availableWidth <= 0 || !options.measureText) {
    return [lineRuns]
  }

  const wrapped: TextLayoutRun[][] = [[]]

  for (const run of lineRuns) {
    const segments = splitRunIntoWrapSegments(run)

    for (const segment of segments) {
      appendSegmentWithWrap(wrapped, segment, options.availableWidth, options.measureText, options.defaultFontSize)
    }
  }

  return wrapped.filter((line) => line.length > 0)
}

function appendSegmentWithWrap(
  wrapped: TextLayoutRun[][],
  segment: TextLayoutRun,
  availableWidth: number,
  measureText: (context: TextMeasureContext) => number,
  defaultFontSize: number,
): void {
  if (segment.text === '') {
    return
  }

  const currentLine = wrapped[wrapped.length - 1]

  if (!currentLine) {
    wrapped.push([segment])
    return
  }

  const nextLine = mergeAdjacentRuns([...currentLine, segment])
  const nextWidth = measureLineWidth(nextLine, measureText, defaultFontSize)

  if (nextWidth <= availableWidth) {
    wrapped[wrapped.length - 1] = nextLine
    return
  }

  const trimmedText = segment.text.trimStart()

  if (trimmedText.length === 0) {
    return
  }

  const trimmedSegment = trimmedText === segment.text ? segment : { ...segment, text: trimmedText }
  const segmentWidth = measureRunWidth(trimmedSegment, measureText, defaultFontSize)

  if (segmentWidth <= availableWidth) {
    wrapped.push([trimmedSegment])
    return
  }

  const brokenSegments = breakSegmentToWidth(trimmedSegment, availableWidth, measureText, defaultFontSize)

  for (const brokenSegment of brokenSegments) {
    wrapped.push([brokenSegment])
  }
}

function splitRunIntoWrapSegments(run: TextLayoutRun): TextLayoutRun[] {
  const segments: TextLayoutRun[] = []
  let current = ''
  let currentKind: 'space' | 'latin' | 'cjk' | 'punctuation' | null = null

  for (const character of normalizeTextRunText(run.text)) {
    const kind = characterKind(character)

    if (kind === 'cjk' || kind === 'punctuation') {
      pushWrapSegment(segments, current, run)
      current = character
      currentKind = kind
      continue
    }

    if (currentKind && currentKind !== kind) {
      pushWrapSegment(segments, current, run)
      current = character
      currentKind = kind
      continue
    }

    current += character
    currentKind = kind
  }

  pushWrapSegment(segments, current, run)
  return mergeCjkPunctuationSegments(segments)
}

function pushWrapSegment(segments: TextLayoutRun[], text: string, run: TextLayoutRun): void {
  if (text.length > 0) {
    segments.push({ text, style: run.style, fontSize: run.fontSize })
  }
}

function characterKind(character: string): 'space' | 'latin' | 'cjk' | 'punctuation' {
  if (/\s/u.test(character)) {
    return 'space'
  }

  if (isCjkCharacter(character)) {
    return 'cjk'
  }

  if (isLineStartForbidden(character) || isLineEndForbidden(character)) {
    return 'punctuation'
  }

  return 'latin'
}

function mergeCjkPunctuationSegments(segments: TextLayoutRun[]): TextLayoutRun[] {
  const merged: TextLayoutRun[] = []

  for (const segment of segments) {
    const previous = merged[merged.length - 1]

    if (isLineStartForbidden(segment.text) && previous) {
      previous.text += segment.text
      continue
    }

    if (previous && isLineEndForbidden(previous.text)) {
      previous.text += segment.text
      continue
    }

    merged.push({ ...segment })
  }

  return merged
}

function isCjkCharacter(character: string): boolean {
  return /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]/u.test(character)
}

function isLineStartForbidden(text: string): boolean {
  return /^[，。！？；：、）】》」』〕〉…,.!?;:%)]/u.test(text)
}

function isLineEndForbidden(text: string): boolean {
  return /[（【《「『〔〈([]$/u.test(text)
}

function breakSegmentToWidth(
  segment: TextLayoutRun,
  availableWidth: number,
  measureText: (context: TextMeasureContext) => number,
  defaultFontSize: number,
): TextLayoutRun[] {
  const width = measureRunWidth(segment, measureText, defaultFontSize)

  if (width <= availableWidth || availableWidth <= 0) {
    return [segment]
  }

  const characters = [...segment.text]
  const chunks: TextLayoutRun[] = []
  let current = ''

  for (const character of characters) {
    const next = `${current}${character}`
    const nextWidth = measureRunWidth({ text: next, style: segment.style }, measureText, defaultFontSize)

    if (current.length > 0 && nextWidth > availableWidth) {
      chunks.push({ text: current, style: segment.style, fontSize: segment.fontSize })
      current = character
      continue
    }

    current = next
  }

  if (current.length > 0) {
    chunks.push({ text: current, style: segment.style, fontSize: segment.fontSize })
  }

  return chunks.length > 0 ? chunks : [segment]
}

function mergeAdjacentRuns(runs: TextLayoutRun[]): TextLayoutRun[] {
  const merged: TextLayoutRun[] = []

  for (const run of runs) {
    const previous = merged[merged.length - 1]

    if (previous && sameTextStyle(previous.style, run.style)) {
      previous.text += run.text
      continue
    }

    merged.push({ ...run })
  }

  return merged
}

function sameTextStyle(left: TextStyle | undefined, right: TextStyle | undefined): boolean {
  return JSON.stringify(left ?? null) === JSON.stringify(right ?? null)
}

function measureLineWidth(
  runs: TextLayoutRun[],
  measureText: (context: TextMeasureContext) => number,
  defaultFontSize: number,
): number {
  return runs.reduce((total, run) => total + measureRunWidth(run, measureText, defaultFontSize), 0)
}

function measureRunWidth(
  run: TextLayoutRun,
  measureText: (context: TextMeasureContext) => number,
  defaultFontSize: number,
): number {
  return measureText({ text: run.text, style: run.style, defaultFontSize })
}

function paragraphBullet(paragraph: Paragraph, autoNumberState: AutoNumberState): string | undefined {
  const bullet = paragraph.style?.bullet

  if (!bullet || bullet.type === 'none') {
    autoNumberState.nextByLevel.clear()
    return undefined
  }

  if (bullet.type === 'character') {
    autoNumberState.nextByLevel.clear()
    return `${bullet.character ?? '•'} `
  }

  const level = paragraph.style?.level ?? 0
  const startAt = bullet.autoNumberStartAt ?? 1
  const current = autoNumberState.nextByLevel.get(level) ?? startAt

  for (const knownLevel of autoNumberState.nextByLevel.keys()) {
    if (knownLevel > level) {
      autoNumberState.nextByLevel.delete(knownLevel)
    }
  }

  autoNumberState.nextByLevel.set(level, current + 1)
  return `${current}. `
}

function bulletStyle(paragraph: Paragraph): TextStyle | undefined {
  const bullet = paragraph.style?.bullet

  if (!bullet) {
    return paragraph.style?.defaultRunStyle
  }

  return {
    ...paragraph.style?.defaultRunStyle,
    ...(bullet.color ? { color: bullet.color } : {}),
    ...(bullet.fontFace ? { fontFace: bullet.fontFace } : {}),
    ...(bullet.fontSize ? { fontSize: bullet.fontSize } : {}),
  }
}

function measureBulletLayout(
  bullet: string | undefined,
  style: TextStyle | undefined,
  defaultFontSize: number,
  measureText: ((context: TextMeasureContext) => number) | undefined,
): BulletLayoutContext | undefined {
  if (!bullet) {
    return undefined
  }

  const width = measureText ? measureText({ text: bullet, style, defaultFontSize }) : 0
  return {
    text: bullet,
    style,
    width,
  }
}

function paragraphAvailableWidth(
  paragraph: Paragraph,
  blockWidth: number,
  lineIndex: number,
  bulletLayout: BulletLayoutContext | undefined,
): number {
  const metrics = paragraphLineMetrics(paragraph, lineIndex, 0, blockWidth, bulletLayout)
  const contentInset = lineIndex === 0 ? metrics.contentLeftX : metrics.leftX
  return Math.max(1, blockWidth - contentInset - metrics.marginRight)
}

function paragraphLineX(
  paragraph: Paragraph,
  lineIndex: number,
  blockX: number,
  blockWidth: number,
  align: TextHorizontalAlign,
  runs: TextLayoutRun[],
  defaultFontSize: number,
  measureText: ((context: TextMeasureContext) => number) | undefined,
  bulletLayout: BulletLayoutContext | undefined,
): number {
  const metrics = paragraphLineMetrics(paragraph, lineIndex, blockX, blockWidth, bulletLayout)

  if (!measureText) {
    if (align === 'center') {
      return metrics.anchorX
    }

    if (align === 'right') {
      return metrics.rightX
    }

    return metrics.leftX
  }

  const lineWidth = measureLineWidth(runs, measureText, defaultFontSize)

  if (align === 'center') {
    return metrics.anchorX - lineWidth / 2
  }

  if (align === 'right') {
    return metrics.rightX - lineWidth
  }

  return lineIndex === 0 ? metrics.contentLeftX - bulletLayoutWidth(bulletLayout) : metrics.leftX
}

function paragraphLineMetrics(
  paragraph: Paragraph,
  lineIndex: number,
  blockX: number,
  blockWidth: number,
  bulletLayout: BulletLayoutContext | undefined,
) {
  const marginLeft = emuToPx(paragraph.style?.marginLeft)
  const marginRight = emuToPx(paragraph.style?.marginRight)
  const indent = emuToPx(paragraph.style?.indent)
  const bulletWidth = bulletLayoutWidth(bulletLayout)
  const firstLineInset = bulletLayout ? Math.max(0, marginLeft + Math.min(indent, 0)) : marginLeft + Math.max(indent, 0)
  const hangingInset = bulletLayout ? Math.max(marginLeft, marginLeft + Math.abs(Math.min(indent, 0))) : marginLeft + Math.abs(Math.min(indent, 0))
  const inset = lineIndex === 0 ? firstLineInset : hangingInset
  const leftX = blockX + inset
  const contentLeftX = lineIndex === 0 ? leftX + bulletWidth : leftX
  const rightX = blockX + blockWidth - marginRight
  const contentAnchorX = blockX + (contentLeftX - marginRight) / 2 + (blockWidth - contentLeftX - marginRight) / 2

  return {
    leftX,
    contentLeftX,
    rightX,
    anchorX: lineIndex === 0 && bulletLayout ? contentAnchorX : blockX + (marginLeft - marginRight) / 2 + blockWidth / 2,
    marginRight,
  }
}

function bulletLayoutWidth(bulletLayout: BulletLayoutContext | undefined): number {
  return bulletLayout?.width ?? 0
}

function paragraphLineHeight(paragraph: Paragraph, defaultFontSize: number): number {
  const percent = paragraph.style?.lineSpacing?.percent

  if (percent) {
    return defaultFontSize * (percent / 100000)
  }

  const points = paragraph.style?.lineSpacing?.points

  if (points) {
    return spacingPointsToPx(points)
  }

  return defaultFontSize * 1.3
}

function paragraphSpaceBefore(paragraph: Paragraph, defaultFontSize: number): number {
  return spacingToPx(paragraph.style?.spaceBefore, defaultFontSize)
}

function paragraphSpaceAfter(paragraph: Paragraph, defaultFontSize: number): number {
  return spacingToPx(paragraph.style?.spaceAfter, defaultFontSize)
}

function paragraphBlockHeight(paragraph: Paragraph, lines: TextLayoutRun[][], defaultFontSize: number): number {
  const lineHeight = paragraphLineHeight(paragraph, defaultFontSize)
  return paragraphSpaceBefore(paragraph, defaultFontSize) + lines.length * lineHeight + paragraphSpaceAfter(paragraph, defaultFontSize)
}

function spacingToPx(spacing: TextSpacing | undefined, defaultFontSize: number): number {
  if (!spacing) {
    return 0
  }

  if (spacing.points !== undefined) {
    return spacingPointsToPx(spacing.points)
  }

  if (spacing.percent !== undefined) {
    return defaultFontSize * (spacing.percent / 100000)
  }

  return 0
}

function paragraphAlign(paragraphs: Paragraph[]): TextHorizontalAlign {
  const align = paragraphs.find((paragraph) => paragraph.style?.align)?.style?.align

  if (align === 'ctr') {
    return 'center'
  }

  if (align === 'r') {
    return 'right'
  }

  return 'left'
}

function baseFontSize(element: TextElement): number {
  const sizes = element.textBody.paragraphs.flatMap((paragraph) => [
    paragraph.style?.defaultRunStyle?.fontSize,
    ...paragraph.runs.map((run) => run.style?.fontSize),
  ])
  const fontSize = sizes.find((value) => typeof value === 'number')
  const fallbackHeight = element.transform?.height ?? 20

  return fontSize ? normalizedFontSize(fontSize) : Math.max(10, Math.min(28, fallbackHeight / 2))
}

function spacingPointsToPx(value: number): number {
  return value / 75
}

function emuToPx(value: number | undefined): number {
  return value ? value / 12700 : 0
}
