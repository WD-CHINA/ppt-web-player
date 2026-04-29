import { parseFill } from '../drawing/Fill'
import { parseLine } from '../drawing/Line'
import { parseTransform } from '../drawing/Transform'
import type { Diagnostic } from '../../diagnostics/Diagnostic'
import { createDiagnostic } from '../../diagnostics/createDiagnostic'
import type {
  ConnectorElement,
  Fill,
  ImageElement,
  ShapeElement,
  SlideElement,
  TextElement,
  UnknownElement,
} from '../../model/Presentation'
import type { PptxPackage } from '../../package/PptxPackage'
import type { XmlNode } from '../../xml/XmlNode'
import * as xml from '../../xml/XmlQuery'

export interface ParsedSlideContent {
  background?: Fill
  elements: SlideElement[]
  diagnostics: Diagnostic[]
}

export async function parseSlide(pptx: PptxPackage, slidePart: string): Promise<ParsedSlideContent> {
  const diagnostics: Diagnostic[] = []
  const root = await pptx.getXml(slidePart)

  if (!root) {
    diagnostics.push(
      createDiagnostic({
        code: 'SLIDE_PART_NOT_FOUND',
        severity: 'warning',
        part: slidePart,
        message: `未能读取 slide part：${slidePart}`,
      }),
    )
    return { elements: [], diagnostics }
  }

  const shapeTree = xml.path(root, ['p:cSld', 'p:spTree'])

  if (!shapeTree) {
    diagnostics.push(
      createDiagnostic({
        code: 'SLIDE_SHAPE_TREE_NOT_FOUND',
        severity: 'warning',
        part: slidePart,
        message: 'slide 中未找到 p:spTree。',
      }),
    )
    return { elements: [], diagnostics }
  }

  const background = parseSlideBackground(root)
  const elements: SlideElement[] = []

  for (const node of xml.children(shapeTree)) {
    await appendShapeTreeChild(pptx, slidePart, node, elements, diagnostics)
  }

  return { background, elements, diagnostics }
}

function parseSlideBackground(root: XmlNode): Fill | undefined {
  const background = xml.path(root, ['p:cSld', 'p:bg'])

  return background ? parseFill(background) : undefined
}

async function appendShapeTreeChild(
  pptx: PptxPackage,
  slidePart: string,
  node: XmlNode,
  elements: SlideElement[],
  diagnostics: Diagnostic[],
): Promise<void> {
  if (node.name === 'p:nvGrpSpPr' || node.name === 'p:grpSpPr') {
    return
  }

  if (node.name === 'p:grpSp') {
    for (const child of xml.children(node)) {
      await appendShapeTreeChild(pptx, slidePart, child, elements, diagnostics)
    }

    return
  }

  const element = await parseShapeTreeChild(pptx, slidePart, node, elements.length, diagnostics)

  if (element) {
    elements.push(element)
  }
}

async function parseShapeTreeChild(
  pptx: PptxPackage,
  slidePart: string,
  node: XmlNode,
  index: number,
  diagnostics: Diagnostic[],
): Promise<SlideElement | null> {
  if (node.name === 'p:sp') {
    return parseShapeElement(node, index)
  }

  if (node.name === 'p:cxnSp') {
    return parseConnectorElement(node, index)
  }

  if (node.name === 'p:pic') {
    return parseImageElement(pptx, slidePart, node, index, diagnostics)
  }

  return createUnknownElement(slidePart, node, index, diagnostics)
}

function parseShapeElement(node: XmlNode, index: number): TextElement | ShapeElement | null {
  const textBody = xml.child(node, 'p:txBody')
  const text = xml.text(textBody).trim()
  const transform = parseTransform(node)
  const fill = parseFill(node)
  const line = parseLine(node)

  if (text) {
    return {
      id: elementId(index),
      index,
      name: shapeName(node),
      transform,
      fill,
      line,
      type: 'text',
      text,
    }
  }

  if (!transform && !fill && !line) {
    return null
  }

  return {
    id: elementId(index),
    index,
    name: shapeName(node),
    transform,
    fill,
    line,
    type: 'shape',
  }
}

function parseConnectorElement(node: XmlNode, index: number): ConnectorElement {
  return {
    id: elementId(index),
    index,
    name: shapeName(node),
    transform: parseTransform(node),
    fill: parseFill(node),
    line: parseLine(node),
    type: 'connector',
  }
}

async function parseImageElement(
  pptx: PptxPackage,
  slidePart: string,
  node: XmlNode,
  index: number,
  diagnostics: Diagnostic[],
): Promise<ImageElement> {
  const relationshipId = findFirstAttribute(node, 'r:embed')

  if (!relationshipId) {
    diagnostics.push(
      createDiagnostic({
        code: 'IMAGE_RELATIONSHIP_NOT_FOUND',
        severity: 'warning',
        part: slidePart,
        message: '图片元素缺少 r:embed。',
      }),
    )

    return {
      id: elementId(index),
      index,
      name: shapeName(node),
      transform: parseTransform(node),
      type: 'image',
      isExternal: false,
    }
  }

  const relationship = await pptx.resolveRelationship(slidePart, relationshipId)

  if (!relationship) {
    return {
      id: elementId(index),
      index,
      name: shapeName(node),
      transform: parseTransform(node),
      type: 'image',
      relationshipId,
      isExternal: false,
    }
  }

  return {
    id: elementId(index),
    index,
    name: shapeName(node),
    transform: parseTransform(node),
    type: 'image',
    relationshipId,
    part: relationship.path,
    isExternal: relationship.isExternal,
  }
}

function createUnknownElement(
  slidePart: string,
  node: XmlNode,
  index: number,
  diagnostics: Diagnostic[],
): UnknownElement {
  diagnostics.push(
    createDiagnostic({
      code: 'UNSUPPORTED_SLIDE_ELEMENT',
      severity: 'info',
      part: slidePart,
      message: `暂不支持的 slide 元素：${node.name}`,
      detail: { nodeName: node.name },
    }),
  )

  return {
    id: elementId(index),
    index,
    name: shapeName(node),
    transform: parseTransform(node),
    line: parseLine(node),
    type: 'unknown',
    nodeName: node.name,
  }
}

function findFirstAttribute(node: XmlNode, name: string): string | undefined {
  return xml.attr(node, name) ?? node.children.map((child) => findFirstAttribute(child, name)).find(Boolean)
}

function shapeName(node: XmlNode): string | undefined {
  return findFirstAttribute(xml.child(node, 'p:nvSpPr') ?? xml.child(node, 'p:nvPicPr') ?? node, 'name')
}

function elementId(index: number): string {
  return `element-${index + 1}`
}
