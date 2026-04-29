import { readFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { parsePptx } from '../parser/parsePptx'

describe('parsePptx', () => {
  it('reads presentation size and slide list from a real pptx fixture', async () => {
    const filePath = fileURLToPath(new URL('../../../app/public/演示文稿1.pptx', import.meta.url))
    const buffer = await readFile(filePath)
    const result = await parsePptx(buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength))

    expect(result.presentation).not.toBeNull()
    expect(result.presentation?.width).toBeGreaterThan(0)
    expect(result.presentation?.height).toBeGreaterThan(0)
    expect(result.presentation?.slides.length).toBeGreaterThan(0)
    expect(result.presentation?.slides[0]?.part).toMatch(/^ppt\/slides\/slide\d+\.xml$/)
    expect(result.presentation?.slides[0]?.elements).toEqual(expect.any(Array))
  })

  it('extracts text and image placeholders from slide content', async () => {
    const zip = new JSZip()
    zip.file(
      '[Content_Types].xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
</Types>`,
    )
    zip.file(
      'ppt/presentation.xml',
      `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldSz cx="9144000" cy="5143500"/>
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
</p:presentation>`,
    )
    zip.file(
      'ppt/_rels/presentation.xml.rels',
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>`,
    )
    zip.file(
      'ppt/slides/slide1.xml',
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="F2F7FF"/></a:solidFill></p:bgPr></p:bg><p:spTree>
    <p:nvGrpSpPr/><p:grpSpPr/>
    <p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="1828800"/><a:ext cx="2743200" cy="914400"/></a:xfrm><a:ln w="38100"><a:solidFill><a:srgbClr val="111827"/></a:solidFill></a:ln></p:spPr><p:txBody><a:p><a:r><a:t>Hello PPTX</a:t></a:r></a:p></p:txBody></p:sp>
    <p:pic><p:nvPicPr><p:cNvPr id="3" name="Picture"/></p:nvPicPr><p:blipFill><a:blip r:embed="rId2"/></p:blipFill><p:spPr><a:xfrm><a:off x="4572000" y="914400"/><a:ext cx="1828800" cy="1828800"/></a:xfrm></p:spPr></p:pic>
    <p:sp><p:nvSpPr><p:cNvPr id="4" name="Accent Box"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="3657600"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:solidFill><a:srgbClr val="3486f7"/></a:solidFill><a:ln w="76200"><a:solidFill><a:srgbClr val="EF4444"/></a:solidFill></a:ln></p:spPr></p:sp>
    <p:cxnSp><p:nvCxnSpPr><p:cNvPr id="5" name="Connector"/></p:nvCxnSpPr><p:spPr><a:xfrm><a:off x="3657600" y="3657600"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:solidFill><a:srgbClr val="FFAA00"/></a:solidFill><a:ln w="114300"><a:solidFill><a:srgbClr val="0F766E"/></a:solidFill></a:ln></p:spPr></p:cxnSp>
    <p:grpSp><p:nvGrpSpPr/><p:grpSpPr/>
      <p:sp><p:nvSpPr><p:cNvPr id="6" name="Grouped Text"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="457200"/><a:ext cx="1828800" cy="457200"/></a:xfrm></p:spPr><p:txBody><a:p><a:r><a:t>Grouped Hello</a:t></a:r></a:p></p:txBody></p:sp>
      <p:sp><p:nvSpPr><p:cNvPr id="7" name="Grouped Box"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="2743200" y="457200"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:solidFill><a:srgbClr val="22C55E"/></a:solidFill></p:spPr></p:sp>
    </p:grpSp>
    <p:graphicFrame/>
  </p:spTree></p:cSld>
</p:sld>`,
    )
    zip.file(
      'ppt/slides/_rels/slide1.xml.rels',
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>`,
    )
    zip.file('ppt/media/image1.png', new Uint8Array([1, 2, 3]))

    const input = await zip.generateAsync({ type: 'arraybuffer' })
    const result = await parsePptx(input)
    const slide = result.presentation?.slides[0]
    const elements = slide?.elements

    expect(slide?.background).toEqual({ type: 'solid', color: '#F2F7FF' })
    expect(elements).toEqual([
      expect.objectContaining({
        type: 'text',
        text: 'Hello PPTX',
        name: 'Title',
        transform: { x: 96, y: 192, width: 288, height: 96 },
        line: { color: '#111827', width: 4 },
      }),
      expect.objectContaining({
        type: 'image',
        relationshipId: 'rId2',
        part: 'ppt/media/image1.png',
        transform: { x: 480, y: 96, width: 192, height: 192 },
      }),
      expect.objectContaining({
        type: 'shape',
        name: 'Accent Box',
        transform: { x: 96, y: 384, width: 192, height: 96 },
        fill: { type: 'solid', color: '#3486F7' },
        line: { color: '#EF4444', width: 8 },
      }),
      expect.objectContaining({
        type: 'connector',
        name: 'Connector',
        transform: { x: 384, y: 384, width: 192, height: 96 },
        fill: { type: 'solid', color: '#FFAA00' },
        line: { color: '#0F766E', width: 12 },
      }),
      expect.objectContaining({
        type: 'text',
        text: 'Grouped Hello',
        name: 'Grouped Text',
        transform: { x: 96, y: 48, width: 192, height: 48 },
      }),
      expect.objectContaining({
        type: 'shape',
        name: 'Grouped Box',
        transform: { x: 288, y: 48, width: 96, height: 96 },
        fill: { type: 'solid', color: '#22C55E' },
      }),
      expect.objectContaining({ type: 'unknown', nodeName: 'p:graphicFrame' }),
    ])
    expect(elements?.map((element) => element.id)).toEqual([
      'element-1',
      'element-2',
      'element-3',
      'element-4',
      'element-5',
      'element-6',
      'element-7',
    ])
    expect(elements?.map((element) => element.index)).toEqual([0, 1, 2, 3, 4, 5, 6])
    expect(elements?.some((element) => element.type === 'unknown' && element.nodeName === 'p:grpSp')).toBe(false)
    expect(result.diagnostics).toEqual(
      expect.arrayContaining([expect.objectContaining({ code: 'UNSUPPORTED_SLIDE_ELEMENT' })]),
    )
  })
})
