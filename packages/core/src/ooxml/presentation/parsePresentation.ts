import type { Diagnostic } from '../../diagnostics/Diagnostic'
import { createDiagnostic } from '../../diagnostics/createDiagnostic'
import type { Presentation } from '../../model/Presentation'
import { findOverrideByContentType } from '../../package/ContentTypes'
import type { PptxPackage } from '../../package/PptxPackage'
import * as xml from '../../xml/XmlQuery'
import { parseSlideList } from './parseSlideList'

const PRESENTATION_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'
const EMU_PER_INCH = 914400
const PX_PER_INCH = 96

export async function parsePresentation(pptx: PptxPackage): Promise<Presentation | null> {
  const diagnostics: Diagnostic[] = []
  const contentTypes = await pptx.getContentTypes()
  const presentationOverride = findOverrideByContentType(contentTypes, PRESENTATION_CONTENT_TYPE)
  const presentationPart = presentationOverride?.partName ?? 'ppt/presentation.xml'
  const root = await pptx.getXml(presentationPart)

  if (!root) {
    diagnostics.push(
      createDiagnostic({
        code: 'PRESENTATION_NOT_FOUND',
        severity: 'error',
        part: presentationPart,
        message: '未找到 presentation.xml。',
      }),
    )
    return null
  }

  const size = parseSlideSize(root, diagnostics, presentationPart)
  const slides = await parseSlideList(pptx, presentationPart, root, diagnostics)

  return {
    id: presentationPart,
    width: size.width,
    height: size.height,
    slides,
    diagnostics,
    metadata: {
      presentationPart,
    },
  }
}

function parseSlideSize(root: NonNullable<Awaited<ReturnType<PptxPackage['getXml']>>>, diagnostics: Diagnostic[], part: string) {
  const sizeNode = xml.child(root, 'p:sldSz')
  const cx = Number(xml.attr(sizeNode, 'cx'))
  const cy = Number(xml.attr(sizeNode, 'cy'))

  if (!Number.isFinite(cx) || !Number.isFinite(cy) || cx <= 0 || cy <= 0) {
    diagnostics.push(
      createDiagnostic({
        code: 'SLIDE_SIZE_NOT_FOUND',
        severity: 'warning',
        part,
        message: '未找到有效的 slide 尺寸，使用 16:9 默认尺寸。',
      }),
    )

    return {
      width: 1280,
      height: 720,
    }
  }

  return {
    width: Math.round((cx / EMU_PER_INCH) * PX_PER_INCH),
    height: Math.round((cy / EMU_PER_INCH) * PX_PER_INCH),
  }
}
