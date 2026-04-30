import type { SlideLayout, SlideMaster } from '../../model/Presentation'
import type { PptxPackage } from '../../package/PptxPackage'

const LAYOUT_RELATIONSHIP_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'
const THEME_RELATIONSHIP_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme'

export async function parseSlideMaster(
  pptx: PptxPackage,
  masterPart: string,
  relationshipId: string,
): Promise<{ master: SlideMaster; layouts: SlideLayout[] }> {
  const relationships = await pptx.getRelationships(masterPart)
  const layouts: SlideLayout[] = []
  let themePart: string | undefined

  for (const relationship of relationships) {
    const resolved = await pptx.resolveRelationship(masterPart, relationship.id)

    if (!resolved) {
      continue
    }

    if (resolved.type === LAYOUT_RELATIONSHIP_TYPE) {
      layouts.push({
        id: relationship.id,
        part: resolved.path,
        relationshipId: relationship.id,
        masterPart,
      })
      continue
    }

    if (resolved.type === THEME_RELATIONSHIP_TYPE) {
      themePart = resolved.path
    }
  }

  return {
    master: {
      id: relationshipId,
      part: masterPart,
      relationshipId,
      themePart,
      layoutParts: layouts.map((layout) => layout.part),
    },
    layouts,
  }
}
