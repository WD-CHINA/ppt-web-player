import type { Diagnostic } from '../../diagnostics/Diagnostic'
import { createDiagnostic } from '../../diagnostics/createDiagnostic'
import type { Slide } from '../../model/Presentation'
import type { PptxPackage } from '../../package/PptxPackage'
import type { XmlNode } from '../../xml/XmlNode'
import * as xml from '../../xml/XmlQuery'
import { parseSlide } from './parseSlide'

const PRESENTATIONML_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'

export async function parseSlideList(
  pptx: PptxPackage,
  presentationPart: string,
  presentationRoot: XmlNode,
  diagnostics: Diagnostic[],
): Promise<Slide[]> {
  const slideIdList = xml.child(presentationRoot, 'p:sldIdLst')

  if (!slideIdList) {
    diagnostics.push(
      createDiagnostic({
        code: 'SLIDE_LIST_NOT_FOUND',
        severity: 'warning',
        part: presentationPart,
        message: 'presentation.xml 中未找到 slide 列表。',
      }),
    )
    return []
  }

  const slides: Slide[] = []

  for (const [index, slideNode] of xml.children(slideIdList, 'p:sldId').entries()) {
    const id = xml.attr(slideNode, 'id') ?? String(index + 1)
    const relationshipId = xml.attr(slideNode, 'r:id')

    if (!relationshipId) {
      diagnostics.push(
        createDiagnostic({
          code: 'RELATIONSHIP_NOT_FOUND',
          severity: 'warning',
          part: presentationPart,
          message: `第 ${index + 1} 页缺少 r:id。`,
        }),
      )
      continue
    }

    const relationship = await pptx.resolveRelationship(presentationPart, relationshipId)

    if (!relationship) {
      continue
    }

    if (relationship.type !== PRESENTATIONML_NS) {
      diagnostics.push(
        createDiagnostic({
          code: 'UNEXPECTED_RELATIONSHIP_TYPE',
          severity: 'info',
          part: presentationPart,
          message: `slide relationship 类型非标准 slide 类型：${relationship.type}`,
        }),
      )
    }

    const slideContent = await parseSlide(pptx, relationship.path)
    diagnostics.push(...slideContent.diagnostics)

    slides.push({
      id,
      index,
      part: relationship.path,
      relationshipId,
      background: slideContent.background,
      elements: slideContent.elements,
      diagnostics: slideContent.diagnostics,
    })
  }

  return slides
}
