import { parseFill } from '../drawing/Fill'
import { DIAGNOSTIC_CODES } from '../../diagnostics/codes'
import { pushDiagnostic, slideDiagnosticContext } from '../../diagnostics/context'
import type { Diagnostic } from '../../diagnostics/Diagnostic'
import type { Fill, PresentationTheme, SlideElement, SlideLayout, SlideMaster } from '../../model/Presentation'
import type { PptxPackage } from '../../package/PptxPackage'
import type { XmlNode } from '../../xml/XmlNode'
import * as xml from '../../xml/XmlQuery'
import { parseImageElement } from './parseImage'
import { createUnknownElement, parseConnectorElement, parseShapeElement } from './parseShape'

export interface ParsedSlideContent {
  background?: Fill
  elements: SlideElement[]
  diagnostics: Diagnostic[]
}

export interface SlideParseContext {
  theme?: PresentationTheme
  layout?: SlideLayout
  master?: SlideMaster
}

export async function parseSlide(
  pptx: PptxPackage,
  slidePart: string,
  slideIndex?: number,
  parseContext: SlideParseContext = {},
): Promise<ParsedSlideContent> {
  const diagnostics: Diagnostic[] = []
  const context = slideDiagnosticContext(slidePart, slideIndex)
  const root = await pptx.getXml(slidePart)

  if (!root) {
    pushDiagnostic(
      diagnostics,
      {
        code: DIAGNOSTIC_CODES.slidePartNotFound,
        severity: 'warning',
        message: `未能读取 slide part：${slidePart}`,
      },
      context,
    )
    return { elements: [], diagnostics }
  }

  const shapeTree = xml.path(root, ['p:cSld', 'p:spTree'])

  if (!shapeTree) {
    pushDiagnostic(
      diagnostics,
      {
        code: DIAGNOSTIC_CODES.slideShapeTreeNotFound,
        severity: 'warning',
        message: 'slide 中未找到 p:spTree。',
      },
      context,
    )
    return { elements: [], diagnostics }
  }

  const background = parseSlideBackground(root, parseContext)
  const elements: SlideElement[] = []

  for (const node of xml.children(shapeTree)) {
    await appendShapeTreeChild(pptx, slidePart, slideIndex, node, elements, diagnostics, parseContext)
  }

  return { background, elements, diagnostics }
}

function parseSlideBackground(root: XmlNode, context: SlideParseContext): Fill | undefined {
  const background = xml.path(root, ['p:cSld', 'p:bg'])

  return background ? parseFill(background, context.theme) : (context.layout?.defaults?.background ?? context.master?.defaults?.background)
}

async function appendShapeTreeChild(
  pptx: PptxPackage,
  slidePart: string,
  slideIndex: number | undefined,
  node: XmlNode,
  elements: SlideElement[],
  diagnostics: Diagnostic[],
  parseContext: SlideParseContext,
): Promise<void> {
  if (node.name === 'p:nvGrpSpPr' || node.name === 'p:grpSpPr') {
    return
  }

  if (node.name === 'p:grpSp') {
    for (const child of xml.children(node)) {
      await appendShapeTreeChild(pptx, slidePart, slideIndex, child, elements, diagnostics, parseContext)
    }

    return
  }

  const element = await parseShapeTreeChild(pptx, slidePart, slideIndex, node, elements.length, diagnostics, parseContext)

  if (element) {
    elements.push(element)
  }
}

async function parseShapeTreeChild(
  pptx: PptxPackage,
  slidePart: string,
  slideIndex: number | undefined,
  node: XmlNode,
  index: number,
  diagnostics: Diagnostic[],
  parseContext: SlideParseContext,
): Promise<SlideElement | null> {
  if (node.name === 'p:sp') {
    return parseShapeElement(node, index, slidePart, slideIndex, diagnostics, parseContext)
  }

  if (node.name === 'p:cxnSp') {
    return parseConnectorElement(node, index, slidePart, parseContext)
  }

  if (node.name === 'p:pic') {
    return parseImageElement(pptx, slidePart, slideIndex, node, index, diagnostics)
  }

  return createUnknownElement(node, index, slidePart, slideIndex, diagnostics)
}
