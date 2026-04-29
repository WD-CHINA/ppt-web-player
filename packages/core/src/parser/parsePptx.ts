import type { Diagnostic } from '../diagnostics/Diagnostic'
import type { Presentation } from '../model/Presentation'
import { PptxPackage } from '../package/PptxPackage'
import { parsePresentation } from '../ooxml/presentation/parsePresentation'

export interface ParsePptxResult {
  presentation: Presentation | null
  diagnostics: Diagnostic[]
}

export async function parsePptx(input: Blob | ArrayBuffer): Promise<ParsePptxResult> {
  const pptx = await PptxPackage.load(input)
  const presentation = await parsePresentation(pptx)

  return {
    presentation,
    diagnostics: [...pptx.getDiagnostics(), ...(presentation?.diagnostics ?? [])],
  }
}
