import { parseTransform } from '../drawing/Transform'
import { DIAGNOSTIC_CODES } from '../../diagnostics/codes'
import { elementDiagnosticContext, pushDiagnostic } from '../../diagnostics/context'
import type { Diagnostic } from '../../diagnostics/Diagnostic'
import type { ImageCrop, ImageElement } from '../../model/Presentation'
import type { PptxPackage } from '../../package/PptxPackage'
import type { XmlNode } from '../../xml/XmlNode'
import { parseImageFill } from './parseImageFill'
import { elementId, shapeName } from './parseShape'

export async function parseImageElement(
  pptx: PptxPackage,
  slidePart: string,
  slideIndex: number | undefined,
  node: XmlNode,
  index: number,
  diagnostics: Diagnostic[],
): Promise<ImageElement> {
  const imageFill = await parseImageFill(pptx, slidePart, node, diagnostics, {
    slideIndex,
    elementId: elementId(index),
    missingMessage: '图片元素缺少 r:embed。',
    externalMessage: '外部图片关系暂不加载。',
  })

  if (!imageFill) {
    const diagnostic = pushDiagnostic(
      diagnostics,
      {
        code: DIAGNOSTIC_CODES.imageRelationshipNotFound,
        severity: 'warning',
        message: '图片元素缺少 r:embed。',
      },
      elementDiagnosticContext(slidePart, slideIndex, elementId(index)),
    )

    return createImageElement(node, index, slidePart, {
      isExternal: false,
      diagnostics: [diagnostic],
    })
  }

  return createImageElement(node, index, slidePart, {
    relationshipId: imageFill.relationshipId,
    imagePart: imageFill.imagePart,
    isExternal: imageFill.isExternal,
    crop: imageFill.crop,
    opacity: imageFill.opacity,
    diagnostics: imageFill.diagnostics,
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
    crop?: ImageCrop
    opacity?: number
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
    opacity: options.opacity ?? 1,
    zIndex: index,
    relationshipId: options.relationshipId,
    imagePart: options.imagePart,
    crop: options.crop,
    image: {
      part: options.imagePart,
      isExternal: options.isExternal,
    },
    isExternal: options.isExternal,
    ...(options.diagnostics ? { diagnostics: options.diagnostics } : {}),
  }
}
