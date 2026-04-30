import { parseFill } from '../drawing/Fill'
import { parseLine } from '../drawing/Line'
import { parseTransform } from '../drawing/Transform'
import { DIAGNOSTIC_CODES } from '../../diagnostics/codes'
import { elementDiagnosticContext, pushDiagnostic } from '../../diagnostics/context'
import type { Diagnostic } from '../../diagnostics/Diagnostic'
import type { ConnectorElement, PresentationTheme, ShapeElement, SlideElement, TextElement, UnknownElement } from '../../model/Presentation'
import type { XmlNode } from '../../xml/XmlNode'
import * as xml from '../../xml/XmlQuery'
import { parseGeometry } from './parseGeometry'
import { parseTextBody } from './parseTextBody'

export function parseShapeElement(node: XmlNode, index: number, slidePart: string, theme?: PresentationTheme): TextElement | ShapeElement | null {
  const textBody = parseTextBody(node, theme)
  const text = textBody?.paragraphs.map((paragraph) => paragraph.text).join('\n').trim() ?? ''
  const base = createElementBase(node, index, slidePart)

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

export function parseConnectorElement(node: XmlNode, index: number, slidePart: string): ConnectorElement {
  return {
    ...createElementBase(node, index, slidePart),
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
    ...createElementBase(node, index, slidePart),
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

function createElementBase(node: XmlNode, index: number, slidePart: string) {
  return {
    id: elementId(index),
    index,
    name: shapeName(node),
    transform: parseTransform(node),
    fill: parseFill(node),
    line: parseLine(node),
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

export function flattenText(element: SlideElement): string | undefined {
  return element.type === 'text' ? element.textBody.paragraphs.map((paragraph) => paragraph.text).join('\n') : undefined
}
