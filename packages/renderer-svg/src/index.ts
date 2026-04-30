import type { Fill, ImageCrop, LineEndStyle, LineStyle, Presentation, Slide, SlideBackground, SlideElement, TextElement, TextStyle, Transform } from '@pptx-player/core'
import { layoutTextElement, normalizeTextRunText, textStyleFontSize } from '@pptx-player/layout'

export const rendererSvgPackageName = '@pptx-player/renderer-svg'

export interface RenderSlideToSvgOptions {
  presentation: Pick<Presentation, 'width' | 'height'>
  slide: Slide
  mediaUrls?: Record<string, string>
  ariaLabel?: string
  className?: string
}

export function renderSlideToSvg(options: RenderSlideToSvgOptions): string {
  const { presentation, slide, mediaUrls = {}, ariaLabel, className } = options
  const label = escapeAttribute(ariaLabel ?? `${slide.part} preview`)
  const classAttr = className ? ` class="${escapeAttribute(className)}"` : ''
  const elements = slide.elements.filter((element) => element.transform)

  return [
    `<svg${classAttr} role="img" aria-label="${label}" viewBox="0 0 ${presentation.width} ${presentation.height}" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink">`,
    renderDefs(),
    renderSlideBackground(slide.background, presentation, mediaUrls),
    ...elements.map((element) => renderElement(element, mediaUrls)),
    '</svg>',
  ].join('')
}

function renderSlideBackground(
  background: SlideBackground | undefined,
  presentation: Pick<Presentation, 'width' | 'height'>,
  mediaUrls: Record<string, string>,
): string {
  if (background?.type === 'image') {
    const href = background.fill.imagePart ? mediaUrls[background.fill.imagePart] : undefined

    if (href) {
      const image = renderImage(
        { x: 0, y: 0, width: presentation.width, height: presentation.height },
        href,
        background.fill.crop,
        'none',
      )

      return background.fill.opacity === undefined || background.fill.opacity >= 1 ? image : `<g opacity="${background.fill.opacity}">${image}</g>`
    }
  }

  const fill = background?.type === 'fill' ? background.fill : undefined
  return `<rect width="100%" height="100%" rx="12" fill="${escapeAttribute(solidFillColor(fill, '#f8fafc'))}" />`
}

function renderDefs(): string {
  return `<defs>
    <marker id="ppt-marker-triangle" markerHeight="8" markerWidth="8" orient="auto" refX="7" refY="4">
      <path d="M 0 0 L 8 4 L 0 8 z" fill="context-stroke" />
    </marker>
    <marker id="ppt-marker-stealth" markerHeight="8" markerWidth="9" orient="auto" refX="8" refY="4">
      <path d="M 0 0 L 9 4 L 0 8 L 3 4 z" fill="context-stroke" />
    </marker>
    <marker id="ppt-marker-diamond" markerHeight="8" markerWidth="10" orient="auto" refX="9" refY="4">
      <path d="M 1 4 L 5 0 L 9 4 L 5 8 z" fill="context-stroke" />
    </marker>
    <marker id="ppt-marker-oval" markerHeight="8" markerWidth="8" orient="auto" refX="7" refY="4">
      <ellipse cx="4" cy="4" rx="4" ry="3" fill="context-stroke" />
    </marker>
  </defs>`
}

function renderElement(element: SlideElement, mediaUrls: Record<string, string>): string {
  if (!element.transform || !element.visible || element.opacity <= 0) {
    return ''
  }

  const content = renderElementContent(element, mediaUrls)

  if (element.opacity >= 1 || !content) {
    return content
  }

  return `<g opacity="${element.opacity}">${content}</g>`
}

function renderElementContent(element: SlideElement, mediaUrls: Record<string, string>): string {
  if (!element.transform) {
    return ''
  }

  if (element.type === 'text') {
    return renderTextElement(element)
  }

  if (element.type === 'image') {
    const href = element.imagePart ? mediaUrls[element.imagePart] : undefined

    if (href) {
      return renderImage(element.transform, href, element.crop, 'xMidYMid meet')
    }

    return [
      '<g>',
      `<rect x="${element.transform.x}" y="${element.transform.y}" width="${element.transform.width}" height="${element.transform.height}" rx="8" fill="#38bdf8" fill-opacity="0.28" stroke="#0284c7" />`,
      `<text x="${element.transform.x + 8}" y="${element.transform.y + 20}" font-size="12" fill="#075985">${escapeText(element.imagePart ?? element.relationshipId ?? 'image')}</text>`,
      '</g>',
    ].join('')
  }

  if (element.type === 'shape' && element.geometry?.preset === 'ellipse') {
    return `<ellipse cx="${element.transform.x + element.transform.width / 2}" cy="${element.transform.y + element.transform.height / 2}" rx="${element.transform.width / 2}" ry="${element.transform.height / 2}" ${renderShapePaint(element.fill, element.line)} />`
  }

  if (element.type === 'shape') {
    const radius = element.geometry?.preset === 'roundRect' ? Math.min(element.transform.width, element.transform.height) * 0.16 : 0
    return `<rect x="${element.transform.x}" y="${element.transform.y}" width="${element.transform.width}" height="${element.transform.height}" rx="${radius}" ry="${radius}" ${renderShapePaint(element.fill, element.line)} />`
  }

  if (element.type === 'connector') {
    return `<line x1="${element.transform.x}" y1="${element.transform.y + element.transform.height / 2}" x2="${element.transform.x + element.transform.width}" y2="${element.transform.y + element.transform.height / 2}" stroke="${escapeAttribute(element.line?.color ?? solidFillColor(element.fill, '#64748b'))}" stroke-width="${element.line?.width ?? 4}" stroke-opacity="${svgStrokeOpacity(element.line)}"${renderOptionalAttribute('stroke-dasharray', svgStrokeDasharray(element.line))}${renderOptionalAttribute('marker-start', svgMarkerStart(element.line))}${renderOptionalAttribute('marker-end', svgMarkerEnd(element.line))} stroke-linecap="round" />`
  }

  return `<rect x="${element.transform.x}" y="${element.transform.y}" width="${element.transform.width}" height="${element.transform.height}" fill="none" stroke="#f97316" stroke-dasharray="8 6" />`
}

function renderImage(transform: Transform | undefined, href: string, crop: ImageCrop | undefined, preserveAspectRatio: string): string {
  if (!transform) {
    return ''
  }

  if (!crop) {
    return `<image x="${transform.x}" y="${transform.y}" width="${transform.width}" height="${transform.height}" href="${escapeAttribute(href)}" preserveAspectRatio="${preserveAspectRatio}" />`
  }

  const sourceX = crop.left
  const sourceY = crop.top
  const sourceWidth = 1 - crop.left - crop.right
  const sourceHeight = 1 - crop.top - crop.bottom

  return [
    `<svg x="${transform.x}" y="${transform.y}" width="${transform.width}" height="${transform.height}" viewBox="${sourceX} ${sourceY} ${sourceWidth} ${sourceHeight}" preserveAspectRatio="none">`,
    `<image x="0" y="0" width="1" height="1" href="${escapeAttribute(href)}" preserveAspectRatio="none" />`,
    '</svg>',
  ].join('')
}

function renderTextElement(element: TextElement): string {
  const layout = layoutTextElement(element, {
    measureText: ({ text, style, defaultFontSize }) => estimateTextWidth(text, style, defaultFontSize),
  })

  if (!layout) {
    return ''
  }

  const content = layout.lines
    .map((line) => {
      const tspans = line.runs.map((run) => renderRunTspan(run.text, run.style, run.fontSize ?? layout.defaultFontSize)).join('')
      return `<tspan x="${line.x}" y="${line.y}">${tspans || escapeText('')}</tspan>`
    })
    .join('')

  return `<text fill="#0f172a">${content}</text>`
}

function estimateTextWidth(text: string, style: TextStyle | undefined, defaultFontSize: number): number {
  const fontSize = textStyleFontSize(style, defaultFontSize)
  const normalizedText = normalizeTextRunText(text)

  return [...normalizedText].reduce((width, character) => {
    if (character === ' ') {
      return width + fontSize * 0.33
    }

    const isAscii = character.charCodeAt(0) <= 0x7f
    return width + fontSize * (isAscii ? 0.56 : 1)
  }, 0)
}

function renderRunTspan(text: string, style: TextStyle | undefined, fontSize: number): string {
  const normalizedText = normalizeTextRunText(text)
  const attributes = [
    renderOptionalAttribute('fill', style?.color),
    renderOptionalAttribute('font-family', style?.fontFace),
    renderOptionalAttribute('font-size', String(fontSize)),
    renderOptionalAttribute('font-weight', style?.bold ? '700' : undefined),
    renderOptionalAttribute('font-style', style?.italic ? 'italic' : undefined),
    renderOptionalAttribute('text-decoration', style?.underline ? 'underline' : undefined),
  ]
    .filter(Boolean)
    .join('')

  return `<tspan${attributes}>${escapeText(normalizedText)}</tspan>`
}

function renderShapePaint(fill: Fill | undefined, line: LineStyle | undefined): string {
  return [
    `fill="${escapeAttribute(svgFill(fill))}"`,
    `fill-opacity="${svgFillOpacity(fill)}"`,
    `stroke="${escapeAttribute(line?.color ?? '#334155')}"`,
    `stroke-width="${line?.width ?? 1}"`,
    `stroke-opacity="${svgStrokeOpacity(line)}"`,
    renderOptionalAttribute('stroke-dasharray', svgStrokeDasharray(line)),
  ]
    .filter(Boolean)
    .join(' ')
}

function renderOptionalAttribute(name: string, value: string | undefined): string {
  return value ? ` ${name}="${escapeAttribute(value)}"` : ''
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

function svgStrokeOpacity(line: LineStyle | undefined): number {
  return line?.opacity ?? 1
}

function svgStrokeDasharray(line: LineStyle | undefined): string | undefined {
  if (!line?.dash) {
    return undefined
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
  const pattern = line.dash ? dashPatterns[line.dash] : undefined

  return pattern?.map((value) => value * unit).join(' ')
}

function svgMarkerStart(line: LineStyle | undefined): string | undefined {
  return svgMarker(line?.headEnd)
}

function svgMarkerEnd(line: LineStyle | undefined): string | undefined {
  return svgMarker(line?.tailEnd)
}

function svgMarker(lineEnd: LineEndStyle | undefined): string | undefined {
  const markerIds: Record<string, string> = {
    arrow: 'ppt-marker-triangle',
    diamond: 'ppt-marker-diamond',
    oval: 'ppt-marker-oval',
    stealth: 'ppt-marker-stealth',
    triangle: 'ppt-marker-triangle',
  }
  const markerId = lineEnd ? markerIds[lineEnd.type] : undefined
  return markerId ? `url(#${markerId})` : undefined
}

function escapeText(value: string): string {
  return value.split('&').join('&amp;').split('<').join('&lt;').split('>').join('&gt;')
}

function escapeAttribute(value: string): string {
  return escapeText(value).split('"').join('&quot;')
}
