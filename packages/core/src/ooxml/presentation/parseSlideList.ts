import { pushDiagnostic, slideDiagnosticContext } from '../../diagnostics/context'
import { DIAGNOSTIC_CODES } from '../../diagnostics/codes'
import type { Diagnostic } from '../../diagnostics/Diagnostic'
import type { Slide } from '../../model/Presentation'
import type { PptxPackage } from '../../package/PptxPackage'
import type { XmlNode } from '../../xml/XmlNode'
import * as xml from '../../xml/XmlQuery'
import { parseSlide } from './parseSlide'
import type { PresentationReferences } from './parsePresentationReferences'

const PRESENTATIONML_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'
const LAYOUT_RELATIONSHIP_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'

export async function parseSlideList(
  pptx: PptxPackage,
  presentationPart: string,
  presentationRoot: XmlNode,
  diagnostics: Diagnostic[],
  references: PresentationReferences,
): Promise<Slide[]> {
  const slideIdList = xml.child(presentationRoot, 'p:sldIdLst')

  if (!slideIdList) {
    pushDiagnostic(
      diagnostics,
      {
        code: DIAGNOSTIC_CODES.slideListNotFound,
        severity: 'warning',
        message: 'presentation.xml 中未找到 slide 列表。',
      },
      slideDiagnosticContext(presentationPart),
    )
    return []
  }

  const slides: Slide[] = []

  for (const [index, slideNode] of xml.children(slideIdList, 'p:sldId').entries()) {
    const id = xml.attr(slideNode, 'id') ?? String(index + 1)
    const relationshipId = xml.attr(slideNode, 'r:id')
    const slideContext = slideDiagnosticContext(presentationPart, index)

    if (!relationshipId) {
      pushDiagnostic(
        diagnostics,
        {
          code: DIAGNOSTIC_CODES.relationshipNotFound,
          severity: 'warning',
          message: `第 ${index + 1} 页缺少 r:id。`,
        },
        slideContext,
      )
      continue
    }

    const relationship = await pptx.resolveRelationship(presentationPart, relationshipId)

    if (!relationship) {
      continue
    }

    if (relationship.type !== PRESENTATIONML_NS) {
      pushDiagnostic(
        diagnostics,
        {
          code: DIAGNOSTIC_CODES.unexpectedRelationshipType,
          severity: 'info',
          message: `slide relationship 类型非标准 slide 类型：${relationship.type}`,
        },
        slideContext,
      )
    }

    const layoutReference = await resolveSlideLayoutReference(pptx, relationship.path)
    const matchedLayout = layoutReference ? references.slideLayouts.find((layout) => layout.part === layoutReference.path) : undefined
    const matchedMaster = matchedLayout?.masterPart
      ? references.slideMasters.find((master) => master.part === matchedLayout.masterPart)
      : undefined
    const slideContent = await parseSlide(pptx, relationship.path, index, {
      theme: references.theme,
      layout: matchedLayout,
      master: matchedMaster,
    })
    diagnostics.push(...slideContent.diagnostics)

    slides.push({
      id,
      index,
      part: relationship.path,
      relationshipId,
      background: slideContent.background,
      layoutPart: matchedLayout?.part ?? layoutReference?.path,
      masterPart: matchedMaster?.part,
      themePart: matchedMaster?.themePart ?? references.theme?.part,
      elements: slideContent.elements,
      diagnostics: slideContent.diagnostics,
    })
  }

  return slides
}

async function resolveSlideLayoutReference(pptx: PptxPackage, slidePart: string) {
  const relationships = await pptx.getRelationships(slidePart)
  const layoutRelationship = relationships.find((relationship) => relationship.type === LAYOUT_RELATIONSHIP_TYPE)

  if (!layoutRelationship) {
    return null
  }

  return pptx.resolveRelationship(slidePart, layoutRelationship.id)
}
