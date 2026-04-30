import { parseFill } from '../drawing/Fill'
import { parseLine } from '../drawing/Line'
import { parseTransform } from '../drawing/Transform'
import { DIAGNOSTIC_CODES } from '../../diagnostics/codes'
import { elementDiagnosticContext, pushDiagnostic } from '../../diagnostics/context'
import type { Diagnostic } from '../../diagnostics/Diagnostic'
import type {
  ConnectorElement,
  Fill,
  LineStyle,
  PlaceholderStyle,
  ShapeElement,
  SlideElement,
  TextElement,
  TextStyleDefaults,
  UnknownElement,
} from '../../model/Presentation'
import type { SlideParseContext } from './parseSlide'
import type { XmlNode } from '../../xml/XmlNode'
import * as xml from '../../xml/XmlQuery'
import { parseGeometry } from './parseGeometry'
import { parsePlaceholderMetadata } from './parsePlaceholderStyle'
import { parseTextBody } from './parseTextBody'

export function parseShapeElement(
  node: XmlNode,
  index: number,
  slidePart: string,
  slideIndex: number | undefined,
  diagnostics: Diagnostic[],
  parseContext: SlideParseContext,
): TextElement | ShapeElement | null {
  const inheritedStyle = resolveInheritedPlaceholderStyle(node, index, slidePart, slideIndex, diagnostics, parseContext)
  const textBody = parseTextBody(node, parseContext.theme, inheritedStyle?.text)
  const text = textBody?.paragraphs.map((paragraph) => paragraph.text).join('\n').trim() ?? ''
  const base = createElementBase(node, index, slidePart, parseContext, inheritedStyle)

  if (textBody && text) {
    return {
      ...base,
      type: 'text',
      text,
      textBody,
    }
  }

  if (!base.transform && !base.fill && !base.line) {
    return null
  }

  return {
    ...base,
    type: 'shape',
  }
}

export function parseConnectorElement(node: XmlNode, index: number, slidePart: string, parseContext: SlideParseContext): ConnectorElement {
  return {
    ...createElementBase(node, index, slidePart, parseContext),
    type: 'connector',
  }
}

export function createUnknownElement(
  node: XmlNode,
  index: number,
  slidePart: string,
  slideIndex: number | undefined,
  diagnostics: Diagnostic[],
): UnknownElement {
  const elementIdValue = elementId(index)
  const diagnostic = pushDiagnostic(
    diagnostics,
    {
      code: DIAGNOSTIC_CODES.unsupportedSlideElement,
      severity: 'info',
      message: `暂不支持的 slide 元素：${node.name}`,
      detail: { nodeName: node.name },
    },
    elementDiagnosticContext(slidePart, slideIndex, elementIdValue),
  )

  return {
    ...createElementBase(node, index, slidePart, {}),
    type: 'unknown',
    nodeName: node.name,
    diagnostics: [diagnostic],
  }
}

export function shapeName(node: XmlNode): string | undefined {
  return findFirstAttribute(xml.child(node, 'p:nvSpPr') ?? xml.child(node, 'p:nvPicPr') ?? xml.child(node, 'p:nvCxnSpPr') ?? node, 'name')
}

export function findFirstAttribute(node: XmlNode, name: string): string | undefined {
  return xml.attr(node, name) ?? node.children.map((child) => findFirstAttribute(child, name)).find(Boolean)
}

export function elementId(index: number): string {
  return `element-${index + 1}`
}

function createElementBase(
  node: XmlNode,
  index: number,
  slidePart: string,
  parseContext: SlideParseContext,
  inheritedStyle?: Pick<PlaceholderStyle, 'fill' | 'line'>,
) {
  const fill = mergeFill(parseFill(node, parseContext.theme), inheritedStyle?.fill)
  const line = mergeLine(parseLine(node, parseContext.theme), inheritedStyle?.line)

  return {
    id: elementId(index),
    index,
    name: shapeName(node),
    transform: parseTransform(node),
    ...(fill ? { fill } : {}),
    ...(line ? { line } : {}),
    geometry: parseGeometry(node),
    slidePart,
    source: {
      part: slidePart,
      nodeName: node.name,
    },
    visible: true,
    opacity: 1,
    zIndex: index,
  } as const
}

function resolveInheritedPlaceholderStyle(
  node: XmlNode,
  index: number,
  slidePart: string,
  slideIndex: number | undefined,
  diagnostics: Diagnostic[],
  parseContext: SlideParseContext,
): PlaceholderStyle | undefined {
  const metadata = parsePlaceholderMetadata(node)

  if (!metadata) {
    return undefined
  }

  const layoutStyle = findPlaceholderStyle(parseContext.layout?.defaults?.placeholders, metadata)
  const masterStyle = findPlaceholderStyle(parseContext.master?.defaults?.placeholders, metadata)
  const style = mergePlaceholderStyle(masterStyle, layoutStyle)

  if (!style) {
    pushDiagnostic(
      diagnostics,
      {
        code: DIAGNOSTIC_CODES.styleInheritanceIncomplete,
        severity: 'info',
        message: '未找到匹配的 layout/master placeholder 样式，已仅使用 slide 直接样式。',
        detail: {
          placeholderType: metadata.type,
          placeholderIndex: metadata.index,
          layoutPart: parseContext.layout?.part,
          masterPart: parseContext.master?.part,
        },
      },
      elementDiagnosticContext(slidePart, slideIndex, elementId(index)),
    )
  }

  return style
}

function findPlaceholderStyle(
  placeholders: PlaceholderStyle[] | undefined,
  metadata: Pick<PlaceholderStyle, 'type' | 'index'>,
): PlaceholderStyle | undefined {
  if (!placeholders) {
    return undefined
  }

  if (metadata.index) {
    const match = placeholders.find((placeholder) => placeholder.index === metadata.index)

    if (match) {
      return match
    }
  }

  return metadata.type ? placeholders.find((placeholder) => placeholder.type === metadata.type) : undefined
}

function mergePlaceholderStyle(
  base: PlaceholderStyle | undefined,
  override: PlaceholderStyle | undefined,
): PlaceholderStyle | undefined {
  if (!base && !override) {
    return undefined
  }

  return {
    ...base,
    ...override,
    text: mergeTextDefaults(base?.text, override?.text),
  }
}

function mergeTextDefaults(base: TextStyleDefaults | undefined, override: TextStyleDefaults | undefined): TextStyleDefaults | undefined {
  if (!base && !override) {
    return undefined
  }

  return {
    paragraphs: {
      ...base?.paragraphs,
      ...override?.paragraphs,
    },
  }
}

function mergeFill(direct: Fill | undefined, inherited: Fill | undefined): Fill | undefined {
  return direct ?? inherited
}

function mergeLine(direct: LineStyle | undefined, inherited: LineStyle | undefined): LineStyle | undefined {
  return direct ?? inherited
}

export function flattenText(element: SlideElement): string | undefined {
  return element.type === 'text' ? element.textBody.paragraphs.map((paragraph) => paragraph.text).join('\n') : undefined
}
