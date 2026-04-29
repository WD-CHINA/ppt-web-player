import type { XmlNode } from './XmlNode'

export function child(node: XmlNode | null | undefined, name: string): XmlNode | null {
  return node?.children.find((item) => item.name === name) ?? null
}

export function children(node: XmlNode | null | undefined, name?: string): XmlNode[] {
  if (!node) {
    return []
  }

  if (!name) {
    return node.children
  }

  return node.children.filter((item) => item.name === name)
}

export function attr(node: XmlNode | null | undefined, name: string): string | undefined {
  return node?.attributes[name]
}

export function text(node: XmlNode | null | undefined): string {
  if (!node) {
    return ''
  }

  return [node.text, ...node.children.map(text)].join('')
}

export function path(node: XmlNode | null | undefined, names: string[]): XmlNode | null {
  return names.reduce<XmlNode | null>((current, name) => child(current, name), node ?? null)
}
