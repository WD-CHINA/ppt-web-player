import type { Fill } from '../../model/Presentation'
import type { XmlNode } from '../../xml/XmlNode'
import * as xml from '../../xml/XmlQuery'

export function parseFill(node: XmlNode): Fill | undefined {
  if (findFirstChild(node, 'a:noFill')) {
    return { type: 'none' }
  }

  const solidFill = findFirstChild(node, 'a:solidFill')
  const srgbColorNode = xml.child(solidFill, 'a:srgbClr')
  const srgbColor = normalizeSrgbColor(xml.attr(srgbColorNode, 'val'))

  if (!srgbColor) {
    return undefined
  }

  const opacity = normalizeAlpha(xml.attr(xml.child(srgbColorNode, 'a:alpha'), 'val'))

  return {
    type: 'solid',
    color: srgbColor,
    ...(opacity === undefined ? {} : { opacity }),
  }
}

export function normalizeSrgbColor(value: string | undefined): string | undefined {
  if (!value || !/^[0-9a-f]{6}$/i.test(value)) {
    return undefined
  }

  return `#${value.toUpperCase()}`
}

export function normalizeAlpha(value: string | undefined): number | undefined {
  if (!value) {
    return undefined
  }

  const alpha = Number(value)

  if (!Number.isFinite(alpha)) {
    return undefined
  }

  return Math.min(Math.max(alpha, 0), 100000) / 100000
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
