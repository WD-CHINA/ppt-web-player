import type { Diagnostic } from '../diagnostics/Diagnostic'
import type { Presentation } from '../model/Presentation'
import { PptxPackage } from '../package/PptxPackage'
import { parsePresentation } from '../ooxml/presentation/parsePresentation'

export interface ParsePptxResult {
  presentation: Presentation | null
  diagnostics: Diagnostic[]
  media: Record<string, Blob>
}

export async function parsePptx(input: Blob | ArrayBuffer): Promise<ParsePptxResult> {
  const pptx = await PptxPackage.load(input)
  const presentation = await parsePresentation(pptx)
  const media = presentation ? await collectMedia(pptx, presentation) : {}

  return {
    presentation,
    diagnostics: [...pptx.getDiagnostics(), ...(presentation?.diagnostics ?? [])],
    media,
  }
}

async function collectMedia(pptx: PptxPackage, presentation: Presentation): Promise<Record<string, Blob>> {
  const media: Record<string, Blob> = {}
  const parts = new Set<string>()

  for (const slide of presentation.slides) {
    if (slide.background?.type === 'image' && slide.background.fill.imagePart && !slide.background.fill.isExternal) {
      parts.add(slide.background.fill.imagePart)
    }

    for (const element of slide.elements) {
      if (element.type === 'image' && element.imagePart && !element.isExternal) {
        parts.add(element.imagePart)
      }
    }
  }

  for (const part of parts) {
    const bytes = await pptx.getPart(part)

    if (bytes) {
      media[part] = new Blob([bytes], { type: mediaMimeType(part) })
    }
  }

  return media
}

function mediaMimeType(part: string): string {
  const extension = part.toLowerCase().split('.').pop()

  if (extension === 'png') {
    return 'image/png'
  }

  if (extension === 'jpg' || extension === 'jpeg') {
    return 'image/jpeg'
  }

  if (extension === 'gif') {
    return 'image/gif'
  }

  if (extension === 'webp') {
    return 'image/webp'
  }

  if (extension === 'svg') {
    return 'image/svg+xml'
  }

  if (extension === 'bmp') {
    return 'image/bmp'
  }

  return 'application/octet-stream'
}
