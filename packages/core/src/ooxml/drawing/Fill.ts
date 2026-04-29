import type { Fill } from '../../model/Presentation'
import type { XmlNode } from '../../xml/XmlNode'
import * as xml from '../../xml/XmlQuery'

export function parseFill(node: XmlNode): Fill | undefined {
  const solidFill = findFirstChild(node, 'a:solidFill')
  const srgbColor = normalizeSrgbColor(xml.attr(xml.child(solidFill, 'a:srgbClr'), 'val'))

  if (!srgbColor) {
    return undefined
  }

  return {
    type: 'solid',
    color: srgbColor,
  }
}

export function normalizeSrgbColor(value: string | undefined): string | undefined {
  if (!value || !/^[0-9a-f]{6}$/i.test(value)) {
    return undefined
  }

  return `#${value.toUpperCase()}`
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
