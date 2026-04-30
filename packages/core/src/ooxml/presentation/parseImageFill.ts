import { DIAGNOSTIC_CODES } from '../../diagnostics/codes'
import { elementDiagnosticContext, pushDiagnostic, slideDiagnosticContext } from '../../diagnostics/context'
import type { Diagnostic } from '../../diagnostics/Diagnostic'
import type { ImageCrop, ImageFill } from '../../model/Presentation'
import type { PptxPackage } from '../../package/PptxPackage'
import type { XmlNode } from '../../xml/XmlNode'
import * as xml from '../../xml/XmlQuery'

export interface ParsedImageFill extends ImageFill {
  diagnostics?: Diagnostic[]
}

export interface ParseImageFillOptions {
  slideIndex?: number
  elementId?: string
  missingMessage?: string
  externalMessage?: string
}

export async function parseImageFill(
  pptx: PptxPackage,
  sourcePart: string,
  node: XmlNode,
  diagnostics: Diagnostic[],
  options: ParseImageFillOptions = {},
): Promise<ParsedImageFill | undefined> {
  const blipFill = findBlipFill(node)
  const blip = blipFill ? xml.child(blipFill, 'a:blip') : null

  if (!blipFill) {
    return undefined
  }

  const relationshipId = xml.attr(blip, 'r:embed')
  const externalRelationshipId = xml.attr(blip, 'r:link')
  const context = options.elementId
    ? elementDiagnosticContext(sourcePart, options.slideIndex, options.elementId)
    : slideDiagnosticContext(sourcePart, options.slideIndex)
  const imageDiagnostics: Diagnostic[] = []
  const crop = parseImageCrop(blipFill, diagnostics, imageDiagnostics, context)
  const opacity = parseImageOpacity(blip, diagnostics, imageDiagnostics, context)

  if (!relationshipId) {
    if (!externalRelationshipId) {
      const diagnostic = pushDiagnostic(
        diagnostics,
        {
          code: DIAGNOSTIC_CODES.imageRelationshipNotFound,
          severity: 'warning',
          message: options.missingMessage ?? '图片填充缺少 r:embed。',
        },
        context,
      )
      imageDiagnostics.push(diagnostic)
      return { isExternal: false, crop, opacity, diagnostics: imageDiagnostics }
    }

    const diagnostic = pushDiagnostic(
      diagnostics,
      {
        code: DIAGNOSTIC_CODES.imageRelationshipNotFound,
        severity: 'warning',
        message: options.externalMessage ?? '外部图片关系暂不加载。',
        detail: { relationshipId: externalRelationshipId, relationshipType: 'external' },
      },
      context,
    )
    imageDiagnostics.push(diagnostic)
    return { relationshipId: externalRelationshipId, isExternal: true, crop, opacity, diagnostics: imageDiagnostics }
  }

  const relationship = await pptx.resolveRelationship(sourcePart, relationshipId)

  if (!relationship) {
    const diagnostic = pushDiagnostic(
      diagnostics,
      {
        code: DIAGNOSTIC_CODES.imageRelationshipResolveFailed,
        severity: 'warning',
        message: `图片填充引用的 relationship 不存在：${relationshipId}`,
        detail: { relationshipId },
      },
      context,
    )
    imageDiagnostics.push(diagnostic)
    return { relationshipId, isExternal: false, crop, opacity, diagnostics: imageDiagnostics }
  }

  return {
    relationshipId,
    imagePart: relationship.path,
    isExternal: relationship.isExternal,
    crop,
    opacity,
    ...(imageDiagnostics.length > 0 ? { diagnostics: imageDiagnostics } : {}),
  }
}

function findBlipFill(node: XmlNode): XmlNode | null {
  if (node.name === 'a:blipFill' || node.name === 'p:blipFill') {
    return node
  }

  return xml.child(node, 'a:blipFill') ?? xml.child(node, 'p:blipFill')
}

function parseImageCrop(
  blipFill: XmlNode,
  diagnostics: Diagnostic[],
  imageDiagnostics: Diagnostic[],
  context: ReturnType<typeof slideDiagnosticContext> | ReturnType<typeof elementDiagnosticContext>,
): ImageCrop | undefined {
  const sourceRect = xml.child(blipFill, 'a:srcRect')

  if (!sourceRect) {
    return undefined
  }

  const crop = {
    left: parseCropValue(xml.attr(sourceRect, 'l')) ?? 0,
    top: parseCropValue(xml.attr(sourceRect, 't')) ?? 0,
    right: parseCropValue(xml.attr(sourceRect, 'r')) ?? 0,
    bottom: parseCropValue(xml.attr(sourceRect, 'b')) ?? 0,
  }
  const rawValues = {
    l: xml.attr(sourceRect, 'l'),
    t: xml.attr(sourceRect, 't'),
    r: xml.attr(sourceRect, 'r'),
    b: xml.attr(sourceRect, 'b'),
  }
  const hasInvalidValue = Object.values(rawValues).some((value) => value !== undefined && parseCropValue(value) === undefined)
  const isEmptyCrop = crop.left + crop.right >= 1 || crop.top + crop.bottom >= 1

  if (hasInvalidValue || isEmptyCrop) {
    const diagnostic = pushDiagnostic(
      diagnostics,
      {
        code: DIAGNOSTIC_CODES.imageCropInvalid,
        severity: 'warning',
        message: '图片裁剪参数无效，已忽略裁剪。',
        detail: rawValues,
      },
      context,
    )
    imageDiagnostics.push(diagnostic)
    return undefined
  }

  return crop.left > 0 || crop.top > 0 || crop.right > 0 || crop.bottom > 0 ? crop : undefined
}

function parseCropValue(value: string | undefined): number | undefined {
  if (value === undefined) {
    return undefined
  }

  const parsed = Number(value)

  if (!Number.isFinite(parsed) || parsed < 0 || parsed > 100000) {
    return undefined
  }

  return parsed / 100000
}

function parseImageOpacity(
  blip: XmlNode | null,
  diagnostics: Diagnostic[],
  imageDiagnostics: Diagnostic[],
  context: ReturnType<typeof slideDiagnosticContext> | ReturnType<typeof elementDiagnosticContext>,
): number | undefined {
  const alphaNode = xml.child(blip, 'a:alphaModFix') ?? xml.child(blip, 'a:alpha')
  const rawAlpha = xml.attr(alphaNode, 'amt') ?? xml.attr(alphaNode, 'val')

  if (rawAlpha === undefined) {
    return undefined
  }

  const alpha = Number(rawAlpha)

  if (!Number.isFinite(alpha) || alpha < 0 || alpha > 100000) {
    const diagnostic = pushDiagnostic(
      diagnostics,
      {
        code: DIAGNOSTIC_CODES.imageAlphaInvalid,
        severity: 'warning',
        message: '图片透明度参数无效，已按不透明处理。',
        detail: { alpha: rawAlpha },
      },
      context,
    )
    imageDiagnostics.push(diagnostic)
    return undefined
  }

  return alpha / 100000
}
