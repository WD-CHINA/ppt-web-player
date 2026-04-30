import { parseTransform } from '../drawing/Transform'
import { DIAGNOSTIC_CODES } from '../../diagnostics/codes'
import { elementDiagnosticContext, pushDiagnostic } from '../../diagnostics/context'
import type { Diagnostic } from '../../diagnostics/Diagnostic'
import type { ImageElement } from '../../model/Presentation'
import type { PptxPackage } from '../../package/PptxPackage'
import type { XmlNode } from '../../xml/XmlNode'
import { elementId, shapeName, findFirstAttribute } from './parseShape'

export async function parseImageElement(
  pptx: PptxPackage,
  slidePart: string,
  slideIndex: number | undefined,
  node: XmlNode,
  index: number,
  diagnostics: Diagnostic[],
): Promise<ImageElement> {
  const relationshipId = findFirstAttribute(node, 'r:embed')
  const context = elementDiagnosticContext(slidePart, slideIndex, elementId(index))

  if (!relationshipId) {
    const diagnostic = pushDiagnostic(
      diagnostics,
      {
        code: DIAGNOSTIC_CODES.imageRelationshipNotFound,
        severity: 'warning',
        message: '图片元素缺少 r:embed。',
      },
      context,
    )

    return createImageElement(node, index, slidePart, {
      relationshipId: undefined,
      imagePart: undefined,
      isExternal: false,
      diagnostics: [diagnostic],
    })
  }

  const relationship = await pptx.resolveRelationship(slidePart, relationshipId)

  if (!relationship) {
    const diagnostic = pushDiagnostic(
      diagnostics,
      {
        code: DIAGNOSTIC_CODES.imageRelationshipResolveFailed,
        severity: 'warning',
        message: `图片元素引用的 relationship 不存在：${relationshipId}`,
        detail: { relationshipId },
      },
      context,
    )

    return createImageElement(node, index, slidePart, {
      relationshipId,
      imagePart: undefined,
      isExternal: false,
      diagnostics: [diagnostic],
    })
  }

  return createImageElement(node, index, slidePart, {
    relationshipId,
    imagePart: relationship.path,
    isExternal: relationship.isExternal,
  })
}

function createImageElement(
  node: XmlNode,
  index: number,
  slidePart: string,
  options: {
    relationshipId?: string
    imagePart?: string
    isExternal: boolean
    diagnostics?: Diagnostic[]
  },
): ImageElement {
  return {
    id: elementId(index),
    index,
    name: shapeName(node),
    transform: parseTransform(node),
    type: 'image',
    slidePart,
    source: {
      part: slidePart,
      nodeName: node.name,
    },
    visible: true,
    opacity: 1,
    zIndex: index,
    relationshipId: options.relationshipId,
    imagePart: options.imagePart,
    image: {
      part: options.imagePart,
      isExternal: options.isExternal,
    },
    isExternal: options.isExternal,
    ...(options.diagnostics ? { diagnostics: options.diagnostics } : {}),
  }
}
