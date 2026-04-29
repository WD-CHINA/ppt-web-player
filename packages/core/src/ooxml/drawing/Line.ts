import type { LineStyle } from '../../model/Presentation'
import type { XmlNode } from '../../xml/XmlNode'
import * as xml from '../../xml/XmlQuery'
import { normalizeSrgbColor } from './Fill'
import { emuToPx } from './Transform'

export function parseLine(node: XmlNode): LineStyle | undefined {
  const line = findFirstChild(node, 'a:ln')

  if (!line) {
    return undefined
  }

  const width = parseLineWidth(xml.attr(line, 'w'))
  const color = normalizeSrgbColor(xml.attr(xml.path(line, ['a:solidFill', 'a:srgbClr']), 'val'))

  if (!color && width === undefined) {
    return undefined
  }

  return { color, width }
}

function parseLineWidth(value: string | undefined): number | undefined {
  if (!value) {
    return undefined
  }

  const width = Number(value)

  if (!Number.isFinite(width) || width <= 0) {
    return undefined
  }

  return Math.max(1, emuToPx(width))
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
