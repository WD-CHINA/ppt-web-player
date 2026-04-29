import { describe, expect, it } from 'vitest'

import { parseXml } from '../xml/XmlParser'
import * as xml from '../xml/XmlQuery'

describe('XmlParser', () => {
  it('normalizes txml output behind a stable query API', () => {
    const root = parseXml('<p:sld><p:cSld name="slide"><a:t>Hello</a:t></p:cSld></p:sld>')

    expect(root?.name).toBe('p:sld')
    expect(xml.attr(xml.child(root, 'p:cSld'), 'name')).toBe('slide')
    expect(xml.text(root)).toBe('Hello')
  })
})
