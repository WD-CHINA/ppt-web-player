import type { ShapeGeometry } from '../../model/Presentation'
import type { XmlNode } from '../../xml/XmlNode'
import * as xml from '../../xml/XmlQuery'

export function parseGeometry(node: XmlNode): ShapeGeometry | undefined {
  const presetGeometry = xml.path(node, ['p:spPr', 'a:prstGeom'])
  const preset = xml.attr(presetGeometry, 'prst')

  return preset ? { type: 'preset', preset } : undefined
}
