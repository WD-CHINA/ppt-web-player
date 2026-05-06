import type { Fill, ImageCrop, LineStyle, Presentation, Slide, SlideBackground, SlideElement, TextElement, TextStyle, Transform } from '@pptx-player/core'
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
  drawSlideBackground(context, slide.background, presentation, mediaBitmaps)

  for (const element of orderedVisibleElements(slide.elements)) {
    renderElement(context, element, mediaBitmaps)
  }

  context.restore()
}

function orderedVisibleElements(elements: SlideElement[]): SlideElement[] {
  return elements
    .filter((element) => element.transform)
    .map((element, order) => ({ element, order }))
    .sort((left, right) => left.element.zIndex - right.element.zIndex || left.order - right.order)
    .map(({ element }) => element)
}

function drawSlideBackground(
  context: CanvasRenderingContext2D,
  background: SlideBackground | undefined,
  presentation: Pick<Presentation, 'width' | 'height'>,
  mediaBitmaps: Record<string, CanvasImageSource>,
): void {
  if (background?.type === 'image') {
    const bitmap = background.fill.imagePart ? mediaBitmaps[background.fill.imagePart] : undefined

    if (bitmap) {
      context.save()
      context.globalAlpha *= background.fill.opacity ?? 1
      drawImage(context, { x: 0, y: 0, width: presentation.width, height: presentation.height }, bitmap, background.fill.crop)
      context.restore()
      return
    }
  }

  const fill = background?.type === 'fill' ? background.fill : undefined

  if (fill?.type === 'none') {
    return
  }

  context.fillStyle = solidFillColor(fill, '#f8fafc')
  context.fillRect(0, 0, presentation.width, presentation.height)
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
      drawImage(context, transform, bitmap, element.crop)
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

function drawImage(context: CanvasRenderingContext2D, transform: Transform | undefined, bitmap: CanvasImageSource, crop: ImageCrop | undefined): void {
  if (!transform || !crop) {
    context.drawImage(bitmap, transform?.x ?? 0, transform?.y ?? 0, transform?.width ?? 0, transform?.height ?? 0)
    return
  }

  const sourceWidth = imageSourceWidth(bitmap)
  const sourceHeight = imageSourceHeight(bitmap)

  if (sourceWidth <= 0 || sourceHeight <= 0) {
    context.drawImage(bitmap, transform.x, transform.y, transform.width, transform.height)
    return
  }

  const sx = sourceWidth * crop.left
  const sy = sourceHeight * crop.top
  const sw = sourceWidth * (1 - crop.left - crop.right)
  const sh = sourceHeight * (1 - crop.top - crop.bottom)

  context.drawImage(bitmap, sx, sy, sw, sh, transform.x, transform.y, transform.width, transform.height)
}

function imageSourceWidth(source: CanvasImageSource): number {
  if ('naturalWidth' in source) {
    return source.naturalWidth
  }

  if ('videoWidth' in source) {
    return source.videoWidth
  }

  if ('displayWidth' in source) {
    return source.displayWidth
  }

  return 'width' in source && typeof source.width === 'number' ? source.width : 0
}

function imageSourceHeight(source: CanvasImageSource): number {
  if ('naturalHeight' in source) {
    return source.naturalHeight
  }

  if ('videoHeight' in source) {
    return source.videoHeight
  }

  if ('displayHeight' in source) {
    return source.displayHeight
  }

  return 'height' in source && typeof source.height === 'number' ? source.height : 0
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
  runs: Array<{ text: string; style?: TextStyle; fontSize?: number }>,
  x: number,
  y: number,
  defaultFontSize: number,
): void {
  const widths = runs.map((run) => measureRunWidth(context, run, defaultFontSize))
  let cursorX = x

  for (const [index, run] of runs.entries()) {
    applyTextStyle(context, run.style, run.fontSize ?? defaultFontSize)
    context.fillText(normalizeTextRunText(run.text), cursorX, y)
    cursorX += widths[index] ?? 0
  }
}

function measureRunWidth(context: CanvasRenderingContext2D, run: { text: string; style?: TextStyle; fontSize?: number }, defaultFontSize: number): number {
  applyTextStyle(context, run.style, run.fontSize ?? defaultFontSize)
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
