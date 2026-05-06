import { readFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { DIAGNOSTIC_CODES } from '../diagnostics/codes'
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
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdMaster1"/></p:sldMasterIdLst>
  <p:sldSz cx="9144000" cy="5143500"/>
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
</p:presentation>`,
    )
    zip.file(
      'ppt/_rels/presentation.xml.rels',
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMaster1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>`,
    )
    zip.file(
      'ppt/slideMasters/slideMaster1.xml',
      `<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld name="Master 1"><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld>
</p:sldMaster>`,
    )
    zip.file(
      'ppt/slideMasters/_rels/slideMaster1.xml.rels',
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdTheme1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
  <Relationship Id="rIdLayout1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`,
    )
    zip.file(
      'ppt/slideLayouts/slideLayout1.xml',
      `<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld name="Layout 1"><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld>
</p:sldLayout>`,
    )
    zip.file(
      'ppt/theme/theme1.xml',
      `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:srgbClr val="000000"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F2937"/></a:dk2>
      <a:lt2><a:srgbClr val="F9FAFB"/></a:lt2>
      <a:accent1><a:srgbClr val="3486F7"/></a:accent1>
      <a:accent2><a:srgbClr val="22C55E"/></a:accent2>
      <a:accent3><a:srgbClr val="A855F7"/></a:accent3>
      <a:accent4><a:srgbClr val="F97316"/></a:accent4>
      <a:accent5><a:srgbClr val="0EA5E9"/></a:accent5>
      <a:accent6><a:srgbClr val="EF4444"/></a:accent6>
      <a:hlink><a:srgbClr val="2563EB"/></a:hlink>
      <a:folHlink><a:srgbClr val="7C3AED"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Aptos Display"/></a:majorFont>
      <a:minorFont><a:latin typeface="Aptos"/></a:minorFont>
    </a:fontScheme>
  </a:themeElements>
</a:theme>`,
    )
    zip.file(
      'ppt/slides/slide1.xml',
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="F2F7FF"/></a:solidFill></p:bgPr></p:bg><p:spTree>
    <p:nvGrpSpPr/><p:grpSpPr/>
    <p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="1828800"/><a:ext cx="2743200" cy="914400"/></a:xfrm><a:ln w="38100"><a:solidFill><a:srgbClr val="111827"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr lIns="91440" tIns="45720" rIns="91440" bIns="45720" wrap="none" anchor="ctr"><a:normAutofit fontScale="70000" lnSpcReduction="20000"/></a:bodyPr><a:p><a:r><a:t>Hello PPTX</a:t></a:r></a:p></p:txBody></p:sp>
    <p:pic><p:nvPicPr><p:cNvPr id="3" name="Picture"/></p:nvPicPr><p:blipFill><a:blip r:embed="rId2"><a:alphaModFix amt="65000"/></a:blip><a:srcRect l="10000" t="20000" r="30000" b="10000"/></p:blipFill><p:spPr><a:xfrm><a:off x="4572000" y="914400"/><a:ext cx="1828800" cy="1828800"/></a:xfrm></p:spPr></p:pic>
    <p:sp><p:nvSpPr><p:cNvPr id="4" name="Accent Box"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="3657600"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="3486f7"><a:alpha val="50000"/></a:srgbClr></a:solidFill><a:ln w="76200"><a:solidFill><a:srgbClr val="EF4444"><a:alpha val="25000"/></a:srgbClr></a:solidFill><a:prstDash val="dash"/></a:ln></p:spPr></p:sp>
    <p:sp><p:nvSpPr><p:cNvPr id="8" name="Ellipse"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="6400800" y="3657600"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="A855F7"/></a:solidFill></p:spPr></p:sp>
    <p:sp><p:nvSpPr><p:cNvPr id="9" name="No Fill Box"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="7315200" y="3657600"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="38100"><a:solidFill><a:srgbClr val="111827"/></a:solidFill></a:ln></p:spPr></p:sp>
    <p:cxnSp><p:nvCxnSpPr><p:cNvPr id="5" name="Connector"/></p:nvCxnSpPr><p:spPr><a:xfrm><a:off x="3657600" y="3657600"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:solidFill><a:srgbClr val="FFAA00"/></a:solidFill><a:ln w="114300"><a:solidFill><a:srgbClr val="0F766E"/></a:solidFill><a:tailEnd type="triangle" w="med" len="med"/></a:ln></p:spPr></p:cxnSp>
    <p:grpSp><p:nvGrpSpPr/><p:grpSpPr><a:xfrm><a:off x="960000" y="960000"/><a:ext cx="3840000" cy="1920000"/><a:chOff x="0" y="0"/><a:chExt cx="1920000" cy="960000"/></a:xfrm></p:grpSpPr>
      <p:sp><p:nvSpPr><p:cNvPr id="6" name="Grouped Text"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="480000" y="240000"/><a:ext cx="480000" cy="240000"/></a:xfrm></p:spPr><p:txBody><a:p><a:r><a:t>Grouped Hello</a:t></a:r></a:p></p:txBody></p:sp>
      <p:sp><p:nvSpPr><p:cNvPr id="7" name="Grouped Box"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="960000" y="240000"/><a:ext cx="480000" cy="480000"/></a:xfrm><a:solidFill><a:srgbClr val="22C55E"/></a:solidFill></p:spPr></p:sp>
    </p:grpSp>
    <p:graphicFrame/>
  </p:spTree></p:cSld>
</p:sld>`,
    )
    zip.file(
      'ppt/slides/_rels/slide1.xml.rels',
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayoutSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>`,
    )
    zip.file('ppt/media/image1.png', new Uint8Array([1, 2, 3]))

    const input = await zip.generateAsync({ type: 'arraybuffer' })
    const result = await parsePptx(input)
    const slide = result.presentation?.slides[0]
    const elements = slide?.elements

    expect(result.presentation?.theme).toEqual({
      part: 'ppt/theme/theme1.xml',
      name: 'Office Theme',
      colorScheme: {
        dark1: '#000000',
        light1: '#FFFFFF',
        dark2: '#1F2937',
        light2: '#F9FAFB',
        accent1: '#3486F7',
        accent2: '#22C55E',
        accent3: '#A855F7',
        accent4: '#F97316',
        accent5: '#0EA5E9',
        accent6: '#EF4444',
        hyperlink: '#2563EB',
        followedHyperlink: '#7C3AED',
      },
      fontScheme: {
        majorLatin: 'Aptos Display',
        minorLatin: 'Aptos',
      },
    })
    expect(result.presentation?.slideMasters).toEqual([
      {
        id: 'rIdMaster1',
        part: 'ppt/slideMasters/slideMaster1.xml',
        relationshipId: 'rIdMaster1',
        themePart: 'ppt/theme/theme1.xml',
        layoutParts: ['ppt/slideLayouts/slideLayout1.xml'],
      },
    ])
    expect(result.presentation?.slideLayouts).toEqual([
      {
        id: 'rIdLayout1',
        part: 'ppt/slideLayouts/slideLayout1.xml',
        relationshipId: 'rIdLayout1',
        masterPart: 'ppt/slideMasters/slideMaster1.xml',
      },
    ])
    expect(slide?.layoutPart).toBe('ppt/slideLayouts/slideLayout1.xml')
    expect(slide?.masterPart).toBe('ppt/slideMasters/slideMaster1.xml')
    expect(slide?.themePart).toBe('ppt/theme/theme1.xml')
    expect(slide?.background).toEqual({ type: 'fill', fill: { type: 'solid', color: '#F2F7FF' } })
    expect(elements).toEqual([
      expect.objectContaining({
        type: 'text',
        text: 'Hello PPTX',
        name: 'Title',
        slidePart: 'ppt/slides/slide1.xml',
        source: { part: 'ppt/slides/slide1.xml', nodeName: 'p:sp' },
        visible: true,
        opacity: 1,
        zIndex: 0,
        transform: { x: 96, y: 192, width: 288, height: 96 },
        line: { color: '#111827', width: 4 },
        textBody: {
          properties: {
            inset: { left: 10, top: 5, right: 10, bottom: 5 },
            wrap: false,
            verticalAnchor: 'middle',
            autoFit: { type: 'normal', fontScale: 70000, lineSpaceReduction: 20000 },
          },
          paragraphs: [
            {
              text: 'Hello PPTX',
              runs: [{ text: 'Hello PPTX' }],
            },
          ],
        },
      }),
      expect.objectContaining({
        type: 'image',
        relationshipId: 'rId2',
        slidePart: 'ppt/slides/slide1.xml',
        imagePart: 'ppt/media/image1.png',
        image: { part: 'ppt/media/image1.png', isExternal: false },
        source: { part: 'ppt/slides/slide1.xml', nodeName: 'p:pic' },
        visible: true,
        opacity: 0.65,
        zIndex: 1,
        crop: { left: 0.1, top: 0.2, right: 0.3, bottom: 0.1 },
        transform: { x: 480, y: 96, width: 192, height: 192 },
      }),
      expect.objectContaining({
        type: 'shape',
        name: 'Accent Box',
        slidePart: 'ppt/slides/slide1.xml',
        source: { part: 'ppt/slides/slide1.xml', nodeName: 'p:sp' },
        visible: true,
        opacity: 1,
        zIndex: 2,
        transform: { x: 96, y: 384, width: 192, height: 96 },
        fill: { type: 'solid', color: '#3486F7', opacity: 0.5 },
        line: { color: '#EF4444', width: 8, opacity: 0.25, dash: 'dash' },
        geometry: { type: 'preset', preset: 'rect' },
      }),
      expect.objectContaining({
        type: 'shape',
        name: 'Ellipse',
        slidePart: 'ppt/slides/slide1.xml',
        source: { part: 'ppt/slides/slide1.xml', nodeName: 'p:sp' },
        visible: true,
        opacity: 1,
        zIndex: 3,
        transform: { x: 672, y: 384, width: 96, height: 96 },
        fill: { type: 'solid', color: '#A855F7' },
        geometry: { type: 'preset', preset: 'ellipse' },
      }),
      expect.objectContaining({
        type: 'shape',
        name: 'No Fill Box',
        slidePart: 'ppt/slides/slide1.xml',
        source: { part: 'ppt/slides/slide1.xml', nodeName: 'p:sp' },
        visible: true,
        opacity: 1,
        zIndex: 4,
        transform: { x: 768, y: 384, width: 96, height: 96 },
        fill: { type: 'none' },
        line: { color: '#111827', width: 4 },
        geometry: { type: 'preset', preset: 'roundRect' },
      }),
      expect.objectContaining({
        type: 'connector',
        name: 'Connector',
        slidePart: 'ppt/slides/slide1.xml',
        source: { part: 'ppt/slides/slide1.xml', nodeName: 'p:cxnSp' },
        visible: true,
        opacity: 1,
        zIndex: 5,
        transform: { x: 384, y: 384, width: 192, height: 96 },
        fill: { type: 'solid', color: '#FFAA00' },
        line: { color: '#0F766E', width: 12, tailEnd: { type: 'triangle', width: 'med', length: 'med' } },
      }),
      expect.objectContaining({
        type: 'text',
        text: 'Grouped Hello',
        name: 'Grouped Text',
        slidePart: 'ppt/slides/slide1.xml',
        source: { part: 'ppt/slides/slide1.xml', nodeName: 'p:sp' },
        visible: true,
        opacity: 1,
        zIndex: 6,
        transform: { x: 201, y: 151, width: 100, height: 50 },
        textBody: {
          paragraphs: [
            {
              text: 'Grouped Hello',
              runs: [{ text: 'Grouped Hello' }],
            },
          ],
        },
      }),
      expect.objectContaining({
        type: 'shape',
        name: 'Grouped Box',
        slidePart: 'ppt/slides/slide1.xml',
        source: { part: 'ppt/slides/slide1.xml', nodeName: 'p:sp' },
        visible: true,
        opacity: 1,
        zIndex: 7,
        transform: { x: 303, y: 151, width: 100, height: 100 },
        fill: { type: 'solid', color: '#22C55E' },
      }),
      expect.objectContaining({
        type: 'unknown',
        nodeName: 'p:graphicFrame',
        slidePart: 'ppt/slides/slide1.xml',
        source: { part: 'ppt/slides/slide1.xml', nodeName: 'p:graphicFrame' },
        visible: true,
        opacity: 1,
        zIndex: 8,
        diagnostics: [
          expect.objectContaining({
            code: DIAGNOSTIC_CODES.unsupportedSlideElement,
            elementId: 'element-9',
          }),
        ],
      }),
    ])
    expect(elements?.map((element) => element.id)).toEqual([
      'element-1',
      'element-2',
      'element-3',
      'element-4',
      'element-5',
      'element-6',
      'element-7',
      'element-8',
      'element-9',
    ])
    expect(elements?.map((element) => element.index)).toEqual([0, 1, 2, 3, 4, 5, 6, 7, 8])
    expect(elements?.some((element) => element.type === 'unknown' && element.nodeName === 'p:grpSp')).toBe(false)
    expect(result.media['ppt/media/image1.png']).toBeInstanceOf(Blob)
    expect(result.media['ppt/media/image1.png']?.type).toBe('image/png')
    expect(result.diagnostics).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          code: DIAGNOSTIC_CODES.unsupportedSlideElement,
          part: 'ppt/slides/slide1.xml',
          slideIndex: 0,
          elementId: 'element-9',
        }),
      ]),
    )
  })

  it('inherits layout and master slide backgrounds when slide has no explicit background', async () => {
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
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdMaster1"/></p:sldMasterIdLst>
  <p:sldSz cx="9144000" cy="5143500"/>
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst>
</p:presentation>`,
    )
    zip.file(
      'ppt/_rels/presentation.xml.rels',
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMaster1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
</Relationships>`,
    )
    zip.file(
      'ppt/slideMasters/slideMaster1.xml',
      `<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld name="Master 1"><p:bg><p:bgPr><a:solidFill><a:srgbClr val="111111"/></a:solidFill></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld>
</p:sldMaster>`,
    )
    zip.file(
      'ppt/slideMasters/_rels/slideMaster1.xml.rels',
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rIdLayout2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout2.xml"/>
</Relationships>`,
    )
    zip.file(
      'ppt/slideLayouts/slideLayout1.xml',
      `<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld name="Layout 1"><p:bg><p:bgPr><a:solidFill><a:srgbClr val="222222"/></a:solidFill></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld>
</p:sldLayout>`,
    )
    zip.file(
      'ppt/slideLayouts/slideLayout2.xml',
      `<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld name="Layout 2"><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld>
</p:sldLayout>`,
    )
    zip.file(
      'ppt/slides/slide1.xml',
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld></p:sld>`,
    )
    zip.file(
      'ppt/slides/slide2.xml',
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld></p:sld>`,
    )
    zip.file(
      'ppt/slides/_rels/slide1.xml.rels',
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayoutSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`,
    )
    zip.file(
      'ppt/slides/_rels/slide2.xml.rels',
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayoutSlide2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout2.xml"/>
</Relationships>`,
    )

    const input = await zip.generateAsync({ type: 'arraybuffer' })
    const result = await parsePptx(input)

    expect(result.presentation?.slides[0]?.background).toEqual({ type: 'fill', fill: { type: 'solid', color: '#222222' } })
    expect(result.presentation?.slides[1]?.background).toEqual({ type: 'fill', fill: { type: 'solid', color: '#111111' } })
  })

  it('parses image slide backgrounds and collects their media', async () => {
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
  <p:cSld><p:bg><p:bgPr><a:blipFill><a:blip r:embed="rIdBg"><a:alphaModFix amt="75000"/></a:blip><a:srcRect l="10000" t="20000" r="30000" b="10000"/></a:blipFill></p:bgPr></p:bg><p:spTree>
    <p:nvGrpSpPr/><p:grpSpPr/>
  </p:spTree></p:cSld>
</p:sld>`,
    )
    zip.file(
      'ppt/slides/_rels/slide1.xml.rels',
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdBg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/background.png"/>
</Relationships>`,
    )
    zip.file('ppt/media/background.png', new Uint8Array([1, 2, 3]))

    const input = await zip.generateAsync({ type: 'arraybuffer' })
    const result = await parsePptx(input)
    const background = result.presentation?.slides[0]?.background

    expect(background).toEqual({
      type: 'image',
      fill: {
        relationshipId: 'rIdBg',
        imagePart: 'ppt/media/background.png',
        isExternal: false,
        crop: { left: 0.1, top: 0.2, right: 0.3, bottom: 0.1 },
        opacity: 0.75,
      },
    })
    expect(result.media['ppt/media/background.png']).toBeInstanceOf(Blob)
    expect(result.media['ppt/media/background.png']?.type).toBe('image/png')
  })

  it('attaches part, slideIndex, and elementId to diagnostics when available', async () => {
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
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdMaster1"/></p:sldMasterIdLst>
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
</p:presentation>`,
    )
    zip.file(
      'ppt/_rels/presentation.xml.rels',
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMaster1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>`,
    )
    zip.file(
      'ppt/slideMasters/slideMaster1.xml',
      `<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld name="Master 1"><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld>
</p:sldMaster>`,
    )
    zip.file(
      'ppt/slides/slide1.xml',
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr/><p:grpSpPr/>
    <p:pic><p:nvPicPr><p:cNvPr id="3" name="Picture"/></p:nvPicPr></p:pic>
    <p:graphicFrame/>
  </p:spTree></p:cSld>
</p:sld>`,
    )

    const input = await zip.generateAsync({ type: 'arraybuffer' })
    const result = await parsePptx(input)

    expect(result.presentation?.width).toBe(1280)
    expect(result.presentation?.height).toBe(720)
    expect(result.diagnostics).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          code: DIAGNOSTIC_CODES.slideSizeNotFound,
          part: 'ppt/presentation.xml',
        }),
        expect.objectContaining({
          code: DIAGNOSTIC_CODES.imageRelationshipNotFound,
          part: 'ppt/slides/slide1.xml',
          slideIndex: 0,
          elementId: 'element-1',
        }),
        expect.objectContaining({
          code: DIAGNOSTIC_CODES.unsupportedSlideElement,
          part: 'ppt/slides/slide1.xml',
          slideIndex: 0,
          elementId: 'element-2',
        }),
      ]),
    )
  })

  it('preserves diagnostics when presentation.xml is missing', async () => {
    const zip = new JSZip()
    zip.file(
      '[Content_Types].xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
</Types>`,
    )

    const input = await zip.generateAsync({ type: 'arraybuffer' })
    const result = await parsePptx(input)

    expect(result.presentation).toBeNull()
    expect(result.diagnostics).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          code: DIAGNOSTIC_CODES.presentationNotFound,
          part: 'ppt/presentation.xml',
        }),
      ]),
    )
  })

  it('adds element context when image relationship cannot be resolved', async () => {
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
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdMaster1"/></p:sldMasterIdLst>
  <p:sldSz cx="9144000" cy="5143500"/>
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
</p:presentation>`,
    )
    zip.file(
      'ppt/_rels/presentation.xml.rels',
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMaster1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>`,
    )
    zip.file(
      'ppt/slideMasters/slideMaster1.xml',
      `<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld name="Master 1"><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld>
</p:sldMaster>`,
    )
    zip.file(
      'ppt/slides/slide1.xml',
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr/><p:grpSpPr/>
    <p:pic><p:nvPicPr><p:cNvPr id="3" name="Broken Picture"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdMissing"/></p:blipFill></p:pic>
  </p:spTree></p:cSld>
</p:sld>`,
    )
    zip.file(
      'ppt/slides/_rels/slide1.xml.rels',
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayoutSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`,
    )

    const input = await zip.generateAsync({ type: 'arraybuffer' })
    const result = await parsePptx(input)

    expect(result.diagnostics).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          code: DIAGNOSTIC_CODES.relationshipNotFound,
          part: 'ppt/slides/slide1.xml',
          message: 'Relationship 不存在：rIdMissing',
        }),
        expect.objectContaining({
          code: DIAGNOSTIC_CODES.imageRelationshipResolveFailed,
          part: 'ppt/slides/slide1.xml',
          slideIndex: 0,
          elementId: 'element-1',
        }),
      ]),
    )
    expect(result.presentation?.slides[0]?.elements[0]).toEqual(
      expect.objectContaining({
        type: 'image',
        relationshipId: 'rIdMissing',
        diagnostics: [
          expect.objectContaining({
            code: DIAGNOSTIC_CODES.imageRelationshipResolveFailed,
            slideIndex: 0,
            elementId: 'element-1',
          }),
        ],
      }),
    )
  })

  it('parses richer text runs and paragraph styles', async () => {
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
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdMaster1"/></p:sldMasterIdLst>
  <p:sldSz cx="9144000" cy="5143500"/>
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
</p:presentation>`,
    )
    zip.file(
      'ppt/_rels/presentation.xml.rels',
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMaster1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>`,
    )
    zip.file(
      'ppt/slideMasters/slideMaster1.xml',
      `<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld name="Master 1"><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld>
</p:sldMaster>`,
    )
    zip.file(
      'ppt/slideMasters/_rels/slideMaster1.xml.rels',
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdTheme1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`,
    )
    zip.file(
      'ppt/theme/theme1.xml',
      `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:srgbClr val="000000"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F2937"/></a:dk2>
      <a:lt2><a:srgbClr val="F9FAFB"/></a:lt2>
      <a:accent1><a:srgbClr val="3486F7"/></a:accent1>
      <a:accent2><a:srgbClr val="22C55E"/></a:accent2>
      <a:accent3><a:srgbClr val="A855F7"/></a:accent3>
      <a:accent4><a:srgbClr val="F97316"/></a:accent4>
      <a:accent5><a:srgbClr val="0EA5E9"/></a:accent5>
      <a:accent6><a:srgbClr val="EF4444"/></a:accent6>
      <a:hlink><a:srgbClr val="2563EB"/></a:hlink>
      <a:folHlink><a:srgbClr val="7C3AED"/></a:folHlink>
    </a:clrScheme>
  </a:themeElements>
</a:theme>`,
    )
    zip.file(
      'ppt/slides/slide1.xml',
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr/><p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Rich Text"/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="1828800"/></a:xfrm></p:spPr>
      <p:txBody>
        <a:bodyPr/>
        <a:p>
          <a:pPr algn="ctr" lvl="1" marL="457200" marR="228600" indent="-228600" defTabSz="914400" rtl="1">
            <a:buFont typeface="Wingdings"/>
            <a:buClr><a:schemeClr val="accent4"/></a:buClr>
            <a:buSzPts val="2200"/>
            <a:buChar char="•"/>
            <a:lnSpc><a:spcPct val="150000"/></a:lnSpc>
            <a:spcBef><a:spcPts val="1200"/></a:spcBef>
            <a:spcAft><a:spcPts val="600"/></a:spcAft>
            <a:defRPr sz="1800" b="1"><a:solidFill><a:schemeClr val="accent2"/></a:solidFill><a:latin typeface="Aptos"/></a:defRPr>
          </a:pPr>
          <a:r>
            <a:rPr b="1" i="1" u="sng" sz="2400"><a:solidFill><a:srgbClr val="FF6600"/></a:solidFill><a:latin typeface="Aptos Display"/></a:rPr>
            <a:t>Hello</a:t>
          </a:r>
          <a:tab/>
          <a:br/>
          <a:fld id="{42}" type="slidenum">
            <a:rPr sz="2000"><a:solidFill><a:srgbClr val="00AA55"/></a:solidFill></a:rPr>
            <a:t>42</a:t>
          </a:fld>
        </a:p>
        <a:p>
          <a:pPr>
            <a:buFont typeface="Courier New"/>
            <a:buClr><a:srgbClr val="3366FF"/></a:buClr>
            <a:buSzPct val="125000"/>
            <a:buAutoNum type="arabicPeriod" startAt="3"/>
          </a:pPr>
          <a:r><a:t>Second line</a:t></a:r>
        </a:p>
        <a:p>
          <a:pPr>
            <a:buFont typeface="Symbol"/>
            <a:buClr><a:schemeClr val="accent6"/></a:buClr>
            <a:buSzTx/>
            <a:buNone/>
          </a:pPr>
          <a:r><a:t>Third line</a:t></a:r>
        </a:p>
      </p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:sld>`,
    )

    const input = await zip.generateAsync({ type: 'arraybuffer' })
    const result = await parsePptx(input)
    const textElement = result.presentation?.slides[0]?.elements[0]

    expect(textElement).toEqual(
      expect.objectContaining({
        type: 'text',
        text: 'Hello\t\n42\nSecond line\nThird line',
        textBody: {
          paragraphs: [
            {
              text: 'Hello\t\n42',
              style: {
                align: 'ctr',
                level: 1,
                marginLeft: 457200,
                marginRight: 228600,
                indent: -228600,
                defaultTabSize: 914400,
                rtl: true,
                bullet: {
                  type: 'character',
                  character: '•',
                  fontFace: 'Wingdings',
                  color: '#F97316',
                  fontSize: 2200,
                },
                lineSpacing: {
                  percent: 150000,
                },
                spaceBefore: {
                  points: 1200,
                },
                spaceAfter: {
                  points: 600,
                },
                defaultRunStyle: {
                  bold: true,
                  fontSize: 1800,
                  color: '#22C55E',
                  fontFace: 'Aptos',
                },
              },
              runs: [
                {
                  text: 'Hello',
                  style: {
                    bold: true,
                    italic: true,
                    underline: 'sng',
                    fontSize: 2400,
                    color: '#FF6600',
                    fontFace: 'Aptos Display',
                  },
                },
                {
                  text: '\t',
                  style: {
                    bold: true,
                    fontSize: 1800,
                    color: '#22C55E',
                    fontFace: 'Aptos',
                  },
                },
                { text: '\n' },
                {
                  text: '42',
                  style: {
                    bold: true,
                    fontSize: 2000,
                    color: '#00AA55',
                    fontFace: 'Aptos',
                  },
                },
              ],
            },
            {
              text: 'Second line',
              style: {
                bullet: {
                  type: 'auto-number',
                  autoNumberScheme: 'arabicPeriod',
                  autoNumberStartAt: 3,
                  fontFace: 'Courier New',
                  color: '#3366FF',
                  fontSize: 125000,
                },
              },
              runs: [{ text: 'Second line' }],
            },
            {
              text: 'Third line',
              style: {
                bullet: {
                  type: 'none',
                  fontFace: 'Symbol',
                  color: '#EF4444',
                  fontSize: 0,
                },
              },
              runs: [{ text: 'Third line' }],
            },
          ],
        },
      }),
    )
  })

  it('applies minimal theme and placeholder style inheritance', async () => {
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
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdMaster1"/></p:sldMasterIdLst>
  <p:sldSz cx="9144000" cy="5143500"/>
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
</p:presentation>`,
    )
    zip.file(
      'ppt/_rels/presentation.xml.rels',
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdMaster1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>`,
    )
    zip.file(
      'ppt/theme/theme1.xml',
      `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:srgbClr val="000000"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F2937"/></a:dk2>
      <a:lt2><a:srgbClr val="F9FAFB"/></a:lt2>
      <a:accent1><a:srgbClr val="3486F7"/></a:accent1>
      <a:accent2><a:srgbClr val="22C55E"/></a:accent2>
      <a:accent3><a:srgbClr val="A855F7"/></a:accent3>
      <a:accent4><a:srgbClr val="F97316"/></a:accent4>
      <a:accent5><a:srgbClr val="0EA5E9"/></a:accent5>
      <a:accent6><a:srgbClr val="EF4444"/></a:accent6>
      <a:hlink><a:srgbClr val="2563EB"/></a:hlink>
      <a:folHlink><a:srgbClr val="7C3AED"/></a:folHlink>
    </a:clrScheme>
  </a:themeElements>
</a:theme>`,
    )
    zip.file(
      'ppt/slideMasters/slideMaster1.xml',
      `<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:bg><p:bgPr><a:solidFill><a:schemeClr val="accent5"/></a:solidFill></p:bgPr></p:bg><p:spTree>
    <p:nvGrpSpPr/><p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Master Body"/><p:nvPr><p:ph type="body" idx="2"/></p:nvPr></p:nvSpPr>
      <p:spPr><a:solidFill><a:schemeClr val="accent4"/></a:solidFill><a:ln w="38100"><a:solidFill><a:schemeClr val="accent6"/></a:solidFill></a:ln></p:spPr>
    </p:sp>
  </p:spTree></p:cSld>
</p:sldMaster>`,
    )
    zip.file(
      'ppt/slideMasters/_rels/slideMaster1.xml.rels',
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdTheme1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
  <Relationship Id="rIdLayout1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`,
    )
    zip.file(
      'ppt/slideLayouts/slideLayout1.xml',
      `<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr/><p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Layout Title"/><p:nvPr><p:ph type="title" idx="1"/></p:nvPr></p:nvSpPr>
      <p:spPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:ln w="76200"><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></a:ln></p:spPr>
      <p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr sz="2800" b="1"><a:solidFill><a:schemeClr val="accent3"/></a:solidFill><a:latin typeface="Aptos"/></a:defRPr></a:lvl1pPr></a:lstStyle></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:sldLayout>`,
    )
    zip.file(
      'ppt/slides/slide1.xml',
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr/><p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Inherited Title"/><p:nvPr><p:ph type="title" idx="1"/></p:nvPr></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm></p:spPr>
      <p:txBody><a:p><a:r><a:t>Inherited text</a:t></a:r></a:p></p:txBody>
    </p:sp>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="3" name="Master Body"/><p:nvPr><p:ph type="body" idx="2"/></p:nvPr></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="914400" y="2286000"/><a:ext cx="3657600" cy="914400"/></a:xfrm></p:spPr>
      <p:txBody><a:p><a:r><a:t>Master fallback</a:t></a:r></a:p></p:txBody>
    </p:sp>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="4" name="Direct Override"/><p:nvPr><p:ph type="title" idx="1"/></p:nvPr></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="914400" y="3657600"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr>
      <p:txBody><a:p><a:r><a:t>Direct override</a:t></a:r></a:p></p:txBody>
    </p:sp>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="5" name="Missing Placeholder"/><p:nvPr><p:ph type="subtitle" idx="9"/></p:nvPr></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="5029200" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm></p:spPr>
      <p:txBody><a:p><a:r><a:t>Missing placeholder</a:t></a:r></a:p></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:sld>`,
    )
    zip.file(
      'ppt/slides/_rels/slide1.xml.rels',
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayoutSlide1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`,
    )

    const input = await zip.generateAsync({ type: 'arraybuffer' })
    const result = await parsePptx(input)
    const slide = result.presentation?.slides[0]
    const [title, body, direct, missing] = slide?.elements ?? []

    expect(slide?.background).toEqual({ type: 'fill', fill: { type: 'solid', color: '#0EA5E9' } })
    expect(title).toEqual(
      expect.objectContaining({
        type: 'text',
        fill: { type: 'solid', color: '#3486F7' },
        line: { color: '#22C55E', width: 8 },
        textBody: {
          paragraphs: [
            {
              text: 'Inherited text',
              runs: [
                {
                  text: 'Inherited text',
                  style: { bold: true, fontSize: 2800, color: '#A855F7', fontFace: 'Aptos' },
                },
              ],
              style: {
                defaultRunStyle: { bold: true, fontSize: 2800, color: '#A855F7', fontFace: 'Aptos' },
              },
            },
          ],
        },
      }),
    )
    expect(body).toEqual(
      expect.objectContaining({
        type: 'text',
        fill: { type: 'solid', color: '#F97316' },
        line: { color: '#EF4444', width: 4 },
      }),
    )
    expect(direct).toEqual(
      expect.objectContaining({
        type: 'text',
        fill: { type: 'solid', color: '#FF0000' },
        line: { color: '#22C55E', width: 8 },
      }),
    )
    expect(missing).toEqual(
      expect.objectContaining({
        type: 'text',
        text: 'Missing placeholder',
      }),
    )
    expect(result.diagnostics).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          code: DIAGNOSTIC_CODES.styleInheritanceIncomplete,
          severity: 'info',
          part: 'ppt/slides/slide1.xml',
          slideIndex: 0,
          elementId: 'element-4',
          detail: expect.objectContaining({
            placeholderType: 'subtitle',
            placeholderIndex: '9',
            layoutPart: 'ppt/slideLayouts/slideLayout1.xml',
            masterPart: 'ppt/slideMasters/slideMaster1.xml',
          }),
        }),
      ]),
    )
  })
})
