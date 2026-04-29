import { parse, type tNode } from 'txml'

import type { XmlNode } from './XmlNode'

type TXmlNode = Omit<tNode, 'attributes' | 'children'> & {
  attributes?: Record<string, string | number | boolean>
  children?: TXmlValue[]
}

type TXmlValue = TXmlNode | string | null | undefined

export function parseXml(source: string): XmlNode | null {
  const parsed = parse(source) as TXmlValue[]
  const root =
    parsed
      .map(toXmlNode)
      .find((node): node is XmlNode => node !== null && !node.name.startsWith('?')) ?? null
  return root
}

function toXmlNode(value: TXmlValue): XmlNode | null {
  if (value === null || value === undefined || typeof value !== 'object') {
    return null
  }

  if (!value.tagName) {
    return null
  }

  const children: XmlNode[] = []
  const textParts: string[] = []

  for (const child of value.children ?? []) {
    if (typeof child === 'string' || typeof child === 'number' || typeof child === 'boolean') {
      textParts.push(String(child))
      continue
    }

    const childNode = toXmlNode(child)
    if (childNode) {
      children.push(childNode)
    }
  }

  return {
    name: value.tagName,
    attributes: normalizeAttributes(value.attributes ?? {}),
    children,
    text: textParts.join(''),
  }
}

function normalizeAttributes(
  attributes: Record<string, string | number | boolean>,
): Record<string, string> {
  return Object.fromEntries(Object.entries(attributes).map(([key, value]) => [key, String(value)]))
}
