import { describe, expect, it } from 'vitest'

import type { ElementSource, Slide, TextElement } from '@pptx-player/core'

import { renderSlideToSvg } from './index'

describe('renderSlideToSvg', () => {
  const source: ElementSource = {
    part: '/ppt/slides/slide1.xml',
    nodeName: 'p:sp',
  }

  it('wraps long text into multiple svg lines', () => {
    const element: TextElement = {
      id: 'text-wrap',
      index: 0,
      type: 'text',
      slidePart: '/ppt/slides/slide-wrap.xml',
      source,
      visible: true,
      opacity: 1,
      zIndex: 0,
      transform: { x: 12, y: 20, width: 72, height: 100 },
      text: 'Alpha Beta Gamma',
      textBody: {
        paragraphs: [
          {
            text: 'Alpha Beta Gamma',
            runs: [{ text: 'Alpha Beta Gamma', style: { fontSize: 1600 } }],
            style: { defaultRunStyle: { fontSize: 1600 } },
          },
        ],
      },
    }

    const slide: Slide = {
      id: 'slide-wrap',
      index: 0,
      part: '/ppt/slides/slide-wrap.xml',
      relationshipId: 'rIdWrap',
      layoutPart: '/ppt/slideLayouts/slideLayout1.xml',
      masterPart: '/ppt/slideMasters/slideMaster1.xml',
      background: undefined,
      elements: [element],
      diagnostics: [],
    }

    const svg = renderSlideToSvg({
      presentation: { width: 960, height: 540 },
      slide,
    })

    expect(svg).toContain('<tspan x="12" y="36"><tspan font-size="16">Alpha </tspan></tspan>')
    expect(svg).toContain('<tspan x="12" y="56.8"><tspan font-size="16">Beta </tspan></tspan>')
    expect(svg).toContain('<tspan x="12" y="77.6"><tspan font-size="16">Gamma</tspan></tspan>')
  })

  it('uses shared left-edge positions for center-aligned wrapped svg text', () => {
    const element: TextElement = {
      id: 'text-center-wrap',
      index: 0,
      type: 'text',
      slidePart: '/ppt/slides/slide-center-wrap.xml',
      source,
      visible: true,
      opacity: 1,
      zIndex: 0,
      transform: { x: 20, y: 30, width: 72, height: 100 },
      text: 'Alpha Beta',
      textBody: {
        paragraphs: [
          {
            text: 'Alpha Beta',
            runs: [{ text: 'Alpha Beta', style: { fontSize: 1600 } }],
            style: { align: 'ctr', defaultRunStyle: { fontSize: 1600 } },
          },
        ],
      },
    }

    const slide: Slide = {
      id: 'slide-center-wrap',
      index: 0,
      part: '/ppt/slides/slide-center-wrap.xml',
      relationshipId: 'rIdCenterWrap',
      layoutPart: '/ppt/slideLayouts/slideLayout1.xml',
      masterPart: '/ppt/slideMasters/slideMaster1.xml',
      background: undefined,
      elements: [element],
      diagnostics: [],
    }

    const svg = renderSlideToSvg({
      presentation: { width: 960, height: 540 },
      slide,
    })

    expect(svg).toContain('<tspan x="30.959999999999997" y="46"><tspan font-size="16">Alpha </tspan></tspan>')
    expect(svg).toContain('<tspan x="38.08" y="66.8"><tspan font-size="16">Beta</tspan></tspan>')
  })

  it('renders textBody paragraphs and run styles into svg tspans', () => {
    const element: TextElement = {
      id: 'text-1',
      index: 0,
      type: 'text',
      name: 'Title',
      slidePart: '/ppt/slides/slide1.xml',
      source,
      visible: true,
      opacity: 1,
      zIndex: 0,
      transform: { x: 24, y: 40, width: 260, height: 120 },
      text: 'Alpha Beta\nGamma',
      textBody: {
        paragraphs: [
          {
            text: 'Alpha Beta',
            runs: [
              {
                text: 'Alpha',
                style: { bold: true, color: '#ef4444', fontFace: 'Aptos', fontSize: 1800 },
              },
              {
                text: ' Beta',
                style: { italic: true, underline: 'sng' },
              },
            ],
            style: {
              bullet: { type: 'character', character: '•', color: '#2563eb' },
              defaultRunStyle: { fontSize: 1800 },
            },
          },
          {
            text: 'Gamma',
            runs: [{ text: 'Gamma' }],
            style: {
              align: 'ctr',
              defaultRunStyle: { fontSize: 1600 },
            },
          },
        ],
      },
    }

    const slide: Slide = {
      id: 'slide-1',
      index: 0,
      part: '/ppt/slides/slide1.xml',
      relationshipId: 'rId1',
      layoutPart: '/ppt/slideLayouts/slideLayout1.xml',
      masterPart: '/ppt/slideMasters/slideMaster1.xml',
      background: { type: 'solid', color: '#ffffff', opacity: 1 },
      elements: [element],
      diagnostics: [],
    }

    const svg = renderSlideToSvg({
      presentation: { width: 960, height: 540 },
      slide,
    })

    expect(svg).toContain('<text fill="#0f172a">')
    expect(svg).toContain('<tspan fill="#2563eb" font-size="18">• </tspan>')
    expect(svg).toContain('<tspan fill="#ef4444" font-family="Aptos" font-size="18" font-weight="700">Alpha</tspan>')
    expect(svg).toContain('<tspan font-size="18" font-style="italic" text-decoration="underline"> Beta</tspan>')
    expect(svg).toContain('<tspan x="128.8" y="81.4"><tspan font-size="18">Gamma</tspan></tspan>')
    expect(svg).toContain('>Gamma</tspan>')
  })

  it('renders auto-number bullets and multiline runs on separate lines', () => {
    const element: TextElement = {
      id: 'text-2',
      index: 0,
      type: 'text',
      slidePart: '/ppt/slides/slide2.xml',
      source: {
        part: '/ppt/slides/slide2.xml',
        nodeName: 'p:sp',
      },
      visible: true,
      opacity: 1,
      zIndex: 0,
      transform: { x: 10, y: 10, width: 240, height: 120 },
      text: 'One\nTwo\nThree',
      textBody: {
        paragraphs: [
          {
            text: 'One\nTwo',
            runs: [{ text: 'One\nTwo' }],
            style: {
              bullet: { type: 'auto-number', autoNumberStartAt: 3 },
              defaultRunStyle: { fontSize: 1400 },
            },
          },
          {
            text: 'Three',
            runs: [{ text: 'Three' }],
            style: {
              bullet: { type: 'auto-number', autoNumberStartAt: 3 },
              defaultRunStyle: { fontSize: 1400 },
            },
          },
        ],
      },
    }

    const slide: Slide = {
      id: 'slide-2',
      index: 1,
      part: '/ppt/slides/slide2.xml',
      relationshipId: 'rId2',
      layoutPart: '/ppt/slideLayouts/slideLayout1.xml',
      masterPart: '/ppt/slideMasters/slideMaster1.xml',
      background: undefined,
      elements: [element],
      diagnostics: [],
    }

    const svg = renderSlideToSvg({
      presentation: { width: 960, height: 540 },
      slide,
    })

    expect(svg).toContain('<tspan font-size="14">3. </tspan>')
    expect(svg).toContain('<tspan font-size="14">4. </tspan>')
    expect(svg).toContain('<tspan font-size="14">One</tspan>')
    expect(svg).toContain('<tspan font-size="14">Three</tspan>')

    const lineCount = [...svg.matchAll(/<tspan x="/g)].length
    expect(lineCount).toBe(3)
  })

  it('keeps parent and child auto-number sequences separate in svg output', () => {
    const element: TextElement = {
      id: 'text-auto-number-levels',
      index: 0,
      type: 'text',
      slidePart: '/ppt/slides/slide-auto-number-levels.xml',
      source,
      visible: true,
      opacity: 1,
      zIndex: 0,
      transform: { x: 10, y: 10, width: 260, height: 180 },
      text: 'Parent one\nChild one\nChild two\nParent two',
      textBody: {
        paragraphs: [
          {
            text: 'Parent one',
            runs: [{ text: 'Parent one' }],
            style: { level: 0, bullet: { type: 'auto-number', autoNumberStartAt: 1 }, defaultRunStyle: { fontSize: 1400 } },
          },
          {
            text: 'Child one',
            runs: [{ text: 'Child one' }],
            style: { level: 1, bullet: { type: 'auto-number', autoNumberStartAt: 1 }, defaultRunStyle: { fontSize: 1400 } },
          },
          {
            text: 'Child two',
            runs: [{ text: 'Child two' }],
            style: { level: 1, bullet: { type: 'auto-number', autoNumberStartAt: 1 }, defaultRunStyle: { fontSize: 1400 } },
          },
          {
            text: 'Parent two',
            runs: [{ text: 'Parent two' }],
            style: { level: 0, bullet: { type: 'auto-number', autoNumberStartAt: 1 }, defaultRunStyle: { fontSize: 1400 } },
          },
        ],
      },
    }

    const slide: Slide = {
      id: 'slide-auto-number-levels',
      index: 3,
      part: '/ppt/slides/slide-auto-number-levels.xml',
      relationshipId: 'rIdAutoNumberLevels',
      layoutPart: '/ppt/slideLayouts/slideLayout1.xml',
      masterPart: '/ppt/slideMasters/slideMaster1.xml',
      background: undefined,
      elements: [element],
      diagnostics: [],
    }

    const svg = renderSlideToSvg({ presentation: { width: 960, height: 540 }, slide })

    expect([...svg.matchAll(/>1\. <\/tspan>/g)].length).toBe(2)
    expect([...svg.matchAll(/>2\. <\/tspan>/g)].length).toBe(2)
  })

  it('skips hidden and transparent svg elements', () => {
    const visibleElement: TextElement = {
      id: 'visible-text',
      index: 0,
      type: 'text',
      slidePart: '/ppt/slides/slide-visible.xml',
      source,
      visible: true,
      opacity: 1,
      zIndex: 0,
      transform: { x: 10, y: 10, width: 200, height: 60 },
      text: 'Visible text',
      textBody: {
        paragraphs: [{ text: 'Visible text', runs: [{ text: 'Visible text' }] }],
      },
    }
    const hiddenElement: TextElement = {
      ...visibleElement,
      id: 'hidden-text',
      visible: false,
      text: 'Hidden text',
      textBody: {
        paragraphs: [{ text: 'Hidden text', runs: [{ text: 'Hidden text' }] }],
      },
    }
    const transparentElement: TextElement = {
      ...visibleElement,
      id: 'transparent-text',
      opacity: 0,
      text: 'Transparent text',
      textBody: {
        paragraphs: [{ text: 'Transparent text', runs: [{ text: 'Transparent text' }] }],
      },
    }
    const slide: Slide = {
      id: 'slide-visibility',
      index: 0,
      part: '/ppt/slides/slide-visibility.xml',
      relationshipId: 'rIdVisibility',
      layoutPart: '/ppt/slideLayouts/slideLayout1.xml',
      masterPart: '/ppt/slideMasters/slideMaster1.xml',
      background: undefined,
      elements: [visibleElement, hiddenElement, transparentElement],
      diagnostics: [],
    }

    const svg = renderSlideToSvg({ presentation: { width: 960, height: 540 }, slide })

    expect(svg).toContain('Visible text')
    expect(svg).not.toContain('Hidden text')
    expect(svg).not.toContain('Transparent text')
  })

  it('wraps semi-transparent svg elements in an opacity group', () => {
    const slide: Slide = {
      id: 'slide-opacity',
      index: 0,
      part: '/ppt/slides/slide-opacity.xml',
      relationshipId: 'rIdOpacity',
      layoutPart: '/ppt/slideLayouts/slideLayout1.xml',
      masterPart: '/ppt/slideMasters/slideMaster1.xml',
      background: undefined,
      elements: [
        {
          id: 'opacity-shape',
          index: 0,
          type: 'shape',
          slidePart: '/ppt/slides/slide-opacity.xml',
          source,
          visible: true,
          opacity: 0.5,
          zIndex: 0,
          transform: { x: 10, y: 10, width: 100, height: 40 },
          fill: { type: 'solid', color: '#22c55e', opacity: 1 },
          diagnostics: [],
        },
      ],
      diagnostics: [],
    }

    const svg = renderSlideToSvg({ presentation: { width: 960, height: 540 }, slide })

    expect(svg).toContain('<g opacity="0.5"><rect x="10" y="10" width="100" height="40"')
    expect(svg).toContain('fill="#22c55e"')
  })

  it('renders connector line end markers in svg output', () => {
    const slide: Slide = {
      id: 'slide-marker',
      index: 0,
      part: '/ppt/slides/slide-marker.xml',
      relationshipId: 'rIdMarker',
      layoutPart: '/ppt/slideLayouts/slideLayout1.xml',
      masterPart: '/ppt/slideMasters/slideMaster1.xml',
      background: undefined,
      elements: [
        {
          id: 'connector-marker',
          index: 0,
          type: 'connector',
          slidePart: '/ppt/slides/slide-marker.xml',
          source,
          visible: true,
          opacity: 1,
          zIndex: 0,
          transform: { x: 20, y: 30, width: 120, height: 20 },
          line: { color: '#0f172a', width: 3, headEnd: { type: 'oval' }, tailEnd: { type: 'triangle' } },
          diagnostics: [],
        },
      ],
      diagnostics: [],
    }

    const svg = renderSlideToSvg({ presentation: { width: 960, height: 540 }, slide })

    expect(svg).toContain('marker-start="url(#ppt-marker-oval)"')
    expect(svg).toContain('marker-end="url(#ppt-marker-triangle)"')
  })

  it('applies paragraph indent and spacing to line positions', () => {
    const element: TextElement = {
      id: 'text-3',
      index: 0,
      type: 'text',
      slidePart: '/ppt/slides/slide3.xml',
      source: {
        part: '/ppt/slides/slide3.xml',
        nodeName: 'p:sp',
      },
      visible: true,
      opacity: 1,
      zIndex: 0,
      transform: { x: 20, y: 20, width: 300, height: 180 },
      text: 'Alpha\nBeta\nGamma',
      textBody: {
        paragraphs: [
          {
            text: 'Alpha\nBeta',
            runs: [{ text: 'Alpha\nBeta' }],
            style: {
              bullet: { type: 'character', character: '•' },
              marginLeft: 457200,
              indent: -228600,
              lineSpacing: { percent: 150000 },
              spaceBefore: { points: 1200 },
              spaceAfter: { points: 600 },
              defaultRunStyle: { fontSize: 1800 },
            },
          },
          {
            text: 'Gamma',
            runs: [{ text: 'Gamma' }],
            style: {
              marginLeft: 228600,
              defaultRunStyle: { fontSize: 1800 },
            },
          },
        ],
      },
    }

    const slide: Slide = {
      id: 'slide-3',
      index: 2,
      part: '/ppt/slides/slide3.xml',
      relationshipId: 'rId3',
      layoutPart: '/ppt/slideLayouts/slideLayout1.xml',
      masterPart: '/ppt/slideMasters/slideMaster1.xml',
      background: undefined,
      elements: [element],
      diagnostics: [],
    }

    const svg = renderSlideToSvg({
      presentation: { width: 960, height: 540 },
      slide,
    })

    expect(svg).toContain('<tspan x="38" y="66">')
    expect(svg).toContain('<tspan x="74" y="93">')
    expect(svg).toContain('<tspan x="38" y="148">')
  })
})
