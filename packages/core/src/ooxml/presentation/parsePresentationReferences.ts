import { pushDiagnostic, slideDiagnosticContext } from '../../diagnostics/context'
import { DIAGNOSTIC_CODES } from '../../diagnostics/codes'
import type { Diagnostic } from '../../diagnostics/Diagnostic'
import type { PresentationTheme, SlideLayout, SlideMaster } from '../../model/Presentation'
import type { PptxPackage } from '../../package/PptxPackage'
import type { XmlNode } from '../../xml/XmlNode'
import * as xml from '../../xml/XmlQuery'
import { parseSlideMaster } from './parseSlideMaster'
import { parseTheme } from './parseTheme'

const MASTER_RELATIONSHIP_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster'

export interface PresentationReferences {
  theme?: PresentationTheme
  slideMasters: SlideMaster[]
  slideLayouts: SlideLayout[]
}

export async function parsePresentationReferences(
  pptx: PptxPackage,
  presentationPart: string,
  presentationRoot: XmlNode,
  diagnostics: Diagnostic[],
): Promise<PresentationReferences> {
  const masterIdList = xml.child(presentationRoot, 'p:sldMasterIdLst')
  const context = slideDiagnosticContext(presentationPart)

  if (!masterIdList) {
    return {
      slideMasters: [],
      slideLayouts: [],
    }
  }

  const slideMasters: SlideMaster[] = []
  const slideLayouts: SlideLayout[] = []
  let theme: PresentationTheme | undefined

  for (const masterNode of xml.children(masterIdList, 'p:sldMasterId')) {
    const relationshipId = xml.attr(masterNode, 'r:id')

    if (!relationshipId) {
      pushDiagnostic(
        diagnostics,
        {
          code: DIAGNOSTIC_CODES.relationshipNotFound,
          severity: 'warning',
          message: 'slide master 缺少 r:id。',
        },
        context,
      )
      continue
    }

    const relationship = await pptx.resolveRelationship(presentationPart, relationshipId)

    if (!relationship) {
      continue
    }

    if (relationship.type !== MASTER_RELATIONSHIP_TYPE) {
      pushDiagnostic(
        diagnostics,
        {
          code: DIAGNOSTIC_CODES.unexpectedRelationshipType,
          severity: 'info',
          message: `slide master relationship 类型非标准 master 类型：${relationship.type}`,
        },
        context,
      )
    }

    const parsedMaster = await parseSlideMaster(pptx, relationship.path, relationshipId, diagnostics)
    slideMasters.push(parsedMaster.master)
    slideLayouts.push(...parsedMaster.layouts)

    if (!theme && parsedMaster.master.themePart) {
      theme = await parseTheme(pptx, parsedMaster.master.themePart)
    }
  }

  return {
    theme,
    slideMasters,
    slideLayouts,
  }
}
