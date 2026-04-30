import type { Fill, LineStyle, Presentation, Slide, SlideElement, TextElement, TextStyle } from '@pptx-player/core'
import { layoutTextElement, normalizeTextRunText, textStyleFontSize } from '@pptx-player/layout'

export const rendererCanvasPackageName = '@pptx-player/renderer-canvas'

export interface RenderSlideToCanvasOptions {
  presentation: Pick<Presentation, 'width' | 'height'>
  slide: Slide
  context: CanvasRenderingContext2D
  mediaBitmaps?: Record<string, CanvasImageSource>
  clear?: boolean
}

export function renderSlideToCanvas(options: RenderSlideToCanvasOptions): void {
  const { presentation, slide, context, mediaBitmaps = {}, clear = true } = options
  const width = presentation.width
  const height = presentation.height

  if (clear) {
    context.clearRect(0, 0, width, height)
  }

  context.save()
  context.fillStyle = solidFillColor(slide.background, '#f8fafc')
  context.fillRect(0, 0, width, height)

  for (const element of slide.elements) {
    renderElement(context, element, mediaBitmaps)
  }

  context.restore()
}

function renderElement(
  context: CanvasRenderingContext2D,
  element: SlideElement,
  mediaBitmaps: Record<string, CanvasImageSource>,
): void {
  const transform = element.transform

  if (!transform || !element.visible || element.opacity <= 0) {
    return
  }

  context.save()
  context.globalAlpha *= element.opacity

  if (element.type === 'text') {
    renderTextElement(context, element)
    context.restore()
    return
  }

  if (element.type === 'image') {
    const bitmap = element.imagePart ? mediaBitmaps[element.imagePart] : undefined

    if (bitmap) {
      context.drawImage(bitmap, transform.x, transform.y, transform.width, transform.height)
    } else {
      context.fillStyle = 'rgba(56, 189, 248, 0.28)'
      context.strokeStyle = '#0284c7'
      context.lineWidth = 1
      fillRoundRect(context, transform.x, transform.y, transform.width, transform.height, 8)
      strokeRoundRect(context, transform.x, transform.y, transform.width, transform.height, 8)
      context.fillStyle = '#075985'
      context.font = '12px sans-serif'
      context.textBaseline = 'top'
      context.fillText(element.imagePart ?? element.relationshipId ?? 'image', transform.x + 8, transform.y + 8)
    }

    context.restore()
    return
  }

  if (element.type === 'shape' && element.geometry?.preset === 'ellipse') {
    context.beginPath()
    context.ellipse(
      transform.x + transform.width / 2,
      transform.y + transform.height / 2,
      transform.width / 2,
      transform.height / 2,
      0,
      0,
      Math.PI * 2,
    )
    applyFill(context, element.fill)
    applyStroke(context, element.line)
    context.restore()
    return
  }

  if (element.type === 'shape') {
    const radius = element.geometry?.preset === 'roundRect' ? Math.min(transform.width, transform.height) * 0.16 : 0
    fillRoundRect(context, transform.x, transform.y, transform.width, transform.height, radius)
    applyFill(context, element.fill)
    strokeRoundRect(context, transform.x, transform.y, transform.width, transform.height, radius)
    applyStroke(context, element.line)
    context.restore()
    return
  }

  if (element.type === 'connector') {
    context.beginPath()
    context.moveTo(transform.x, transform.y + transform.height / 2)
    context.lineTo(transform.x + transform.width, transform.y + transform.height / 2)
    applyStroke(context, element.line, solidFillColor(element.fill, '#64748b'), 4)
    context.restore()
    return
  }

  context.strokeStyle = '#f97316'
  context.lineWidth = 1
  context.setLineDash([8, 6])
  context.strokeRect(transform.x, transform.y, transform.width, transform.height)
  context.restore()
}

function renderTextElement(context: CanvasRenderingContext2D, element: TextElement): void {
  const layout = layoutTextElement(element, {
    measureText: ({ text, style, defaultFontSize }) => {
      applyTextStyle(context, style, defaultFontSize)
      return context.measureText(normalizeTextRunText(text)).width
    },
  })

  if (!layout) {
    return
  }

  context.fillStyle = '#0f172a'
  context.textBaseline = 'alphabetic'
  context.textAlign = 'left'

  for (const line of layout.lines) {
    renderTextRuns(context, line.runs, line.x, line.y, layout.defaultFontSize)
  }
}

function renderTextRuns(
  context: CanvasRenderingContext2D,
  runs: Array<{ text: string; style?: TextStyle }>,
  x: number,
  y: number,
  defaultFontSize: number,
): void {
  const widths = runs.map((run) => measureRunWidth(context, run, defaultFontSize))
  let cursorX = x

  for (const [index, run] of runs.entries()) {
    applyTextStyle(context, run.style, defaultFontSize)
    context.fillText(normalizeTextRunText(run.text), cursorX, y)
    cursorX += widths[index] ?? 0
  }
}

function measureRunWidth(context: CanvasRenderingContext2D, run: { text: string; style?: TextStyle }, defaultFontSize: number): number {
  applyTextStyle(context, run.style, defaultFontSize)
  return context.measureText(normalizeTextRunText(run.text)).width
}

function applyTextStyle(context: CanvasRenderingContext2D, style: TextStyle | undefined, defaultFontSize: number): void {
  const fontSize = textStyleFontSize(style, defaultFontSize)
  const fontWeight = style?.bold ? '700' : '400'
  const fontStyle = style?.italic ? 'italic' : 'normal'
  const fontFamily = style?.fontFace ?? 'sans-serif'
  context.font = `${fontStyle} ${fontWeight} ${fontSize}px ${fontFamily}`
  context.fillStyle = style?.color ?? '#0f172a'
}

function applyFill(context: CanvasRenderingContext2D, fill: Fill | undefined): void {
  const fillColor = svgFill(fill)

  if (fillColor === 'none') {
    return
  }

  context.save()
  context.fillStyle = fillColor
  context.globalAlpha *= svgFillOpacity(fill)
  context.fill()
  context.restore()
}

function applyStroke(
  context: CanvasRenderingContext2D,
  line: LineStyle | undefined,
  fallbackColor = '#334155',
  fallbackWidth = 1,
): void {
  context.save()
  context.strokeStyle = line?.color ?? fallbackColor
  context.lineWidth = line?.width ?? fallbackWidth
  context.globalAlpha *= line?.opacity ?? 1
  context.setLineDash(canvasLineDash(line))
  context.stroke()
  context.restore()
}

function canvasLineDash(line: LineStyle | undefined): number[] {
  if (!line?.dash) {
    return []
  }

  const unit = Math.max(1, line.width ?? 1)
  const dashPatterns: Record<string, number[]> = {
    dash: [3, 2],
    dashDot: [3, 2, 1, 2],
    dot: [1, 2],
    lgDash: [6, 2],
    lgDashDot: [6, 2, 1, 2],
    lgDashDotDot: [6, 2, 1, 2, 1, 2],
    sysDash: [3, 1],
    sysDashDot: [3, 1, 1, 1],
    sysDashDotDot: [3, 1, 1, 1, 1, 1],
    sysDot: [1, 1],
  }

  return dashPatterns[line.dash]?.map((value) => value * unit) ?? []
}

function fillRoundRect(
  context: CanvasRenderingContext2D,
  x: number,
  y: number,
  width: number,
  height: number,
  radius: number,
): void {
  context.beginPath()
  roundRectPath(context, x, y, width, height, radius)
}

function strokeRoundRect(
  context: CanvasRenderingContext2D,
  x: number,
  y: number,
  width: number,
  height: number,
  radius: number,
): void {
  context.beginPath()
  roundRectPath(context, x, y, width, height, radius)
}

function roundRectPath(
  context: CanvasRenderingContext2D,
  x: number,
  y: number,
  width: number,
  height: number,
  radius: number,
): void {
  const safeRadius = Math.max(0, Math.min(radius, width / 2, height / 2))
  context.moveTo(x + safeRadius, y)
  context.arcTo(x + width, y, x + width, y + height, safeRadius)
  context.arcTo(x + width, y + height, x, y + height, safeRadius)
  context.arcTo(x, y + height, x, y, safeRadius)
  context.arcTo(x, y, x + width, y, safeRadius)
  context.closePath()
}

function svgFill(fill: Fill | undefined): string {
  if (!fill) {
    return '#cbd5e1'
  }

  return fill.type === 'solid' ? fill.color : 'none'
}

function solidFillColor(fill: Fill | undefined, fallback: string): string {
  return fill?.type === 'solid' ? fill.color : fallback
}

function svgFillOpacity(fill: Fill | undefined): number {
  return fill?.type === 'solid' ? (fill.opacity ?? 1) : 1
}
