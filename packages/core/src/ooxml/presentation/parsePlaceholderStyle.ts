import { parseFill } from '../drawing/Fill'
import { parseLine } from '../drawing/Line'
import type { PlaceholderStyle, PresentationTheme, SlideStyleDefaults } from '../../model/Presentation'
import type { XmlNode } from '../../xml/XmlNode'
import * as xml from '../../xml/XmlQuery'
import { parseTextBodyDefaults } from './parseTextBody'

export function parseSlideStyleDefaults(root: XmlNode | null | undefined, theme?: PresentationTheme): SlideStyleDefaults | undefined {
  if (!root) {
    return undefined
  }

  const background = parseBackground(root, theme)
  const placeholders = xml
    .children(xml.path(root, ['p:cSld', 'p:spTree']), 'p:sp')
    .map((shape) => parsePlaceholderStyle(shape, theme))
    .filter((placeholder): placeholder is PlaceholderStyle => Boolean(placeholder))

  if (!background && placeholders.length === 0) {
    return undefined
  }

  return {
    ...(background ? { background } : {}),
    placeholders,
  }
}

function parseBackground(root: XmlNode, theme?: PresentationTheme) {
  const background = xml.path(root, ['p:cSld', 'p:bg'])

  return background ? parseFill(background, theme) : undefined
}

export function parsePlaceholderMetadata(node: XmlNode): Pick<PlaceholderStyle, 'type' | 'index'> | undefined {
  const placeholder = xml.path(node, ['p:nvSpPr', 'p:nvPr', 'p:ph']) ?? xml.path(node, ['p:nvCxnSpPr', 'p:nvPr', 'p:ph'])
  const type = xml.attr(placeholder, 'type')
  const index = xml.attr(placeholder, 'idx')

  if (!type && !index) {
    return undefined
  }

  return {
    ...(type ? { type } : {}),
    ...(index ? { index } : {}),
  }
}

function parsePlaceholderStyle(node: XmlNode, theme?: PresentationTheme): PlaceholderStyle | undefined {
  const metadata = parsePlaceholderMetadata(node)

  if (!metadata) {
    return undefined
  }

  const fill = parseFill(node, theme)
  const line = parseLine(node, theme)
  const text = parseTextBodyDefaults(node, theme)
  const name = placeholderName(node)

  return {
    ...metadata,
    ...(name ? { name } : {}),
    ...(fill ? { fill } : {}),
    ...(line ? { line } : {}),
    ...(text ? { text } : {}),
  }
}

function placeholderName(node: XmlNode): string | undefined {
  return findFirstAttribute(xml.child(node, 'p:nvSpPr') ?? node, 'name')
}

function findFirstAttribute(node: XmlNode, name: string): string | undefined {
  return xml.attr(node, name) ?? node.children.map((child) => findFirstAttribute(child, name)).find(Boolean)
}
