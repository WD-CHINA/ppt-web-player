import type { Fill, LineEndStyle, LineStyle, SlideElement } from '../model/Presentation'

/**
 * 按 zIndex 排序并按可见性过滤元素。
 *
 * 同时保留原始数组顺序作为次排序键，确保 zIndex 相同时行为稳定。
 */
export function orderedVisibleElements(elements: SlideElement[]): SlideElement[] {
  return elements
    .filter((element) => element.transform)
    .map((element, order) => ({ element, order }))
    .sort((left, right) => left.element.zIndex - right.element.zIndex || left.order - right.order)
    .map(({ element }) => element)
}

/**
 * 从 Fill 中提取 solid 颜色值，不存在时回退到 fallback。
 */
export function solidFillColor(fill: Fill | undefined, fallback: string): string {
  return fill?.type === 'solid' ? fill.color : fallback
}

/**
 * 将 Fill 转为 SVG/Canvas 填充颜色字符串。
 * solid fill 返回颜色值，其他类型返回 'none'。
 */
export function svgFill(fill: Fill | undefined): string {
  if (!fill) {
    return '#cbd5e1'
  }

  return fill.type === 'solid' ? fill.color : 'none'
}

/**
 * 从 Fill 中提取透明度。
 */
export function svgFillOpacity(fill: Fill | undefined): number {
  return fill?.type === 'solid' ? (fill.opacity ?? 1) : 1
}

/**
 * 描边透明度。
 */
export function svgStrokeOpacity(line: LineStyle | undefined): number {
  return line?.opacity ?? 1
}

/**
 * PPTX dash 样式 → SVG/Canvas dasharray 值表。
 */
const DASH_PATTERNS: Record<string, number[]> = {
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

/**
 * 将 PPTX 线条 dash 样式转为 SVG stroke-dasharray 字符串。
 */
export function svgStrokeDasharray(line: LineStyle | undefined): string | undefined {
  if (!line?.dash) {
    return undefined
  }

  const unit = Math.max(1, line.width ?? 1)
  const pattern = DASH_PATTERNS[line.dash]
  return pattern?.map((value) => value * unit).join(' ')
}

/**
 * 将 PPTX 线条 dash 样式转为 Canvas setLineDash 数值数组。
 */
export function canvasLineDash(line: LineStyle | undefined): number[] {
  if (!line?.dash) {
    return []
  }

  const unit = Math.max(1, line.width ?? 1)
  return DASH_PATTERNS[line.dash]?.map((value) => value * unit) ?? []
}

const MARKER_IDS: Record<string, string> = {
  arrow: 'ppt-marker-triangle',
  diamond: 'ppt-marker-diamond',
  oval: 'ppt-marker-oval',
  stealth: 'ppt-marker-stealth',
  triangle: 'ppt-marker-triangle',
}

/**
 * 线条起点装饰标记。
 */
export function svgMarkerStart(line: LineStyle | undefined): string | undefined {
  return svgMarker(line?.headEnd)
}

/**
 * 线条终点装饰标记。
 */
export function svgMarkerEnd(line: LineStyle | undefined): string | undefined {
  return svgMarker(line?.tailEnd)
}

function svgMarker(lineEnd: LineEndStyle | undefined): string | undefined {
  const markerId = lineEnd ? MARKER_IDS[lineEnd.type] : undefined
  return markerId ? `url(#${markerId})` : undefined
}
