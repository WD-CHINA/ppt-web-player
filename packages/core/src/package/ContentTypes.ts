import type { XmlNode } from '../xml/XmlNode'
import * as xml from '../xml/XmlQuery'

export interface ContentTypeOverride {
  partName: string
  contentType: string
}

export interface ContentTypes {
  overrides: ContentTypeOverride[]
}

export function parseContentTypes(root: XmlNode | null): ContentTypes {
  if (!root) {
    return { overrides: [] }
  }

  return {
    overrides: xml.children(root, 'Override').flatMap((node) => {
      const partName = xml.attr(node, 'PartName')
      const contentType = xml.attr(node, 'ContentType')

      if (!partName || !contentType) {
        return []
      }

      return [{ partName: partName.replace(/^\//, ''), contentType }]
    }),
  }
}

export function findOverrideByContentType(contentTypes: ContentTypes, contentType: string): ContentTypeOverride | null {
  return contentTypes.overrides.find((override) => override.contentType === contentType) ?? null
}
