import type { Fill, PresentationTheme } from '../../model/Presentation'
import type { XmlNode } from '../../xml/XmlNode'
import * as xml from '../../xml/XmlQuery'
import { normalizeAlpha, parseDrawingColor } from './Color'

export { normalizeAlpha, normalizeSrgbColor } from './Color'

export function parseFill(node: XmlNode, theme?: PresentationTheme): Fill | undefined {
  if (findFirstChild(node, 'a:noFill')) {
    return { type: 'none' }
  }

  const solidFill = findFirstChild(node, 'a:solidFill')
  const color = parseDrawingColor(solidFill, theme)

  if (!color) {
    return undefined
  }

  const colorNode = xml.child(solidFill, 'a:srgbClr') ?? xml.child(solidFill, 'a:schemeClr')
  const opacity = normalizeAlpha(xml.attr(xml.child(colorNode, 'a:alpha'), 'val'))

  return {
    type: 'solid',
    color,
    ...(opacity === undefined ? {} : { opacity }),
  }
}

function findFirstChild(node: XmlNode, name: string): XmlNode | null {
  if (node.name === name) {
    return node
  }

  for (const child of node.children) {
    const match = findFirstChild(child, name)

    if (match) {
      return match
    }
  }

  return null
}
