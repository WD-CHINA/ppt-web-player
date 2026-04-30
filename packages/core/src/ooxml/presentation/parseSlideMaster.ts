import type { PresentationTheme, SlideLayout, SlideMaster } from '../../model/Presentation'
import type { PptxPackage } from '../../package/PptxPackage'
import * as xml from '../../xml/XmlQuery'
import { parseSlideStyleDefaults } from './parsePlaceholderStyle'
import { parseTheme } from './parseTheme'

const LAYOUT_RELATIONSHIP_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'
const THEME_RELATIONSHIP_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme'

export async function parseSlideMaster(
  pptx: PptxPackage,
  masterPart: string,
  relationshipId: string,
  theme?: PresentationTheme,
): Promise<{ master: SlideMaster; layouts: SlideLayout[] }> {
  const relationships = await pptx.getRelationships(masterPart)
  const layoutReferences: Array<{ relationshipId: string; part: string }> = []
  let themePart: string | undefined

  for (const relationship of relationships) {
    const resolved = await pptx.resolveRelationship(masterPart, relationship.id)

    if (!resolved) {
      continue
    }

    if (resolved.type === LAYOUT_RELATIONSHIP_TYPE) {
      layoutReferences.push({ relationshipId: relationship.id, part: resolved.path })
      continue
    }

    if (resolved.type === THEME_RELATIONSHIP_TYPE) {
      themePart = resolved.path
    }
  }

  const resolvedTheme = theme ?? (themePart ? await parseTheme(pptx, themePart) : undefined)
  const masterRoot = await pptx.getXml(masterPart)
  const masterDefaults = parseSlideStyleDefaults(masterRoot, resolvedTheme)
  const layouts = await Promise.all(layoutReferences.map((layout) => parseSlideLayout(pptx, layout, masterPart, resolvedTheme)))

  return {
    master: {
      id: relationshipId,
      part: masterPart,
      relationshipId,
      themePart,
      layoutParts: layouts.map((layout) => layout.part),
      ...(masterDefaults ? { defaults: masterDefaults } : {}),
    },
    layouts,
  }
}

async function parseSlideLayout(
  pptx: PptxPackage,
  layout: { relationshipId: string; part: string },
  masterPart: string,
  theme?: PresentationTheme,
): Promise<SlideLayout> {
  const root = await pptx.getXml(layout.part)
  const defaults = parseSlideStyleDefaults(root, theme)
  const name = xml.attr(root, 'name')

  return {
    id: layout.relationshipId,
    part: layout.part,
    relationshipId: layout.relationshipId,
    masterPart,
    ...(name ? { name } : {}),
    ...(defaults ? { defaults } : {}),
  }
}
