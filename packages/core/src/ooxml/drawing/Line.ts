import type { LineEndStyle, LineStyle } from '../../model/Presentation'
import type { XmlNode } from '../../xml/XmlNode'
import * as xml from '../../xml/XmlQuery'
import { normalizeAlpha, normalizeSrgbColor } from './Fill'
import { emuToPx } from './Transform'

export function parseLine(node: XmlNode): LineStyle | undefined {
  const line = findFirstChild(node, 'a:ln')

  if (!line) {
    return undefined
  }

  const width = parseLineWidth(xml.attr(line, 'w'))
  const srgbColorNode = xml.path(line, ['a:solidFill', 'a:srgbClr'])
  const color = normalizeSrgbColor(xml.attr(srgbColorNode, 'val'))
  const opacity = normalizeAlpha(xml.attr(xml.child(srgbColorNode, 'a:alpha'), 'val'))
  const dash = parsePresetDash(xml.attr(xml.child(line, 'a:prstDash'), 'val'))
  const headEnd = parseLineEnd(xml.child(line, 'a:headEnd'))
  const tailEnd = parseLineEnd(xml.child(line, 'a:tailEnd'))

  if (
    !color &&
    width === undefined &&
    opacity === undefined &&
    dash === undefined &&
    headEnd === undefined &&
    tailEnd === undefined
  ) {
    return undefined
  }

  return {
    ...(color === undefined ? {} : { color }),
    ...(width === undefined ? {} : { width }),
    ...(opacity === undefined ? {} : { opacity }),
    ...(dash === undefined ? {} : { dash }),
    ...(headEnd === undefined ? {} : { headEnd }),
    ...(tailEnd === undefined ? {} : { tailEnd }),
  }
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

function parsePresetDash(value: string | undefined): string | undefined {
  if (!value || value === 'solid') {
    return undefined
  }

  return value
}

function parseLineEnd(node: XmlNode | null): LineEndStyle | undefined {
  const type = xml.attr(node, 'type')

  if (!type || type === 'none') {
    return undefined
  }

  const width = xml.attr(node, 'w')
  const length = xml.attr(node, 'len')

  return {
    type,
    ...(width === undefined ? {} : { width }),
    ...(length === undefined ? {} : { length }),
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
