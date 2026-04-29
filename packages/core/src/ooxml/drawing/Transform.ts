import type { Transform } from '../../model/Presentation'
import type { XmlNode } from '../../xml/XmlNode'
import * as xml from '../../xml/XmlQuery'

const EMU_PER_INCH = 914400
const PX_PER_INCH = 96

export function parseTransform(node: XmlNode): Transform | undefined {
  const transform = findFirstChild(node, 'a:xfrm')
  const offset = xml.child(transform, 'a:off')
  const extent = xml.child(transform, 'a:ext')

  const x = Number(xml.attr(offset, 'x'))
  const y = Number(xml.attr(offset, 'y'))
  const width = Number(xml.attr(extent, 'cx'))
  const height = Number(xml.attr(extent, 'cy'))

  if (![x, y, width, height].every(Number.isFinite) || width <= 0 || height <= 0) {
    return undefined
  }

  return {
    x: emuToPx(x),
    y: emuToPx(y),
    width: emuToPx(width),
    height: emuToPx(height),
  }
}

export function emuToPx(value: number): number {
  return Math.round((value / EMU_PER_INCH) * PX_PER_INCH)
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
