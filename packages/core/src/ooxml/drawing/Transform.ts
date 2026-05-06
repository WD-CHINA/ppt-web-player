import type { Transform } from '../../model/Presentation'
import type { XmlNode } from '../../xml/XmlNode'
import * as xml from '../../xml/XmlQuery'

const EMU_PER_INCH = 914400
const PX_PER_INCH = 96

export interface GroupTransform {
  frame: Transform
  childFrame: Transform
}

export function parseTransform(node: XmlNode): Transform | undefined {
  const transform = findFirstChild(node, 'a:xfrm')
  return parseTransformFrame(transform, 'a:off', 'a:ext')
}

export function parseGroupTransform(node: XmlNode): GroupTransform | undefined {
  const transform = findFirstChild(node, 'a:xfrm')
  const frame = parseTransformFrame(transform, 'a:off', 'a:ext')
  const childFrame = parseTransformFrame(transform, 'a:chOff', 'a:chExt')

  return frame && childFrame ? { frame, childFrame } : undefined
}

export function mapTransformToGroup(transform: Transform | undefined, group: GroupTransform | undefined): Transform | undefined {
  if (!transform || !group) {
    return transform
  }

  const scaleX = group.frame.width / group.childFrame.width
  const scaleY = group.frame.height / group.childFrame.height

  return {
    x: Math.round(group.frame.x + (transform.x - group.childFrame.x) * scaleX),
    y: Math.round(group.frame.y + (transform.y - group.childFrame.y) * scaleY),
    width: Math.round(transform.width * scaleX),
    height: Math.round(transform.height * scaleY),
  }
}

function parseTransformFrame(node: XmlNode | null, offsetName: string, extentName: string): Transform | undefined {
  const offset = xml.child(node, offsetName)
  const extent = xml.child(node, extentName)

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
