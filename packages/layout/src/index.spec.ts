import { describe, expect, it } from 'vitest'

import type { ElementSource, TextElement } from '@pptx-player/core'

import { layoutTextElement, normalizeTextRunText, textStyleFontSize } from './index'

describe('layoutTextElement', () => {
  const source: ElementSource = {
    part: '/ppt/slides/slide1.xml',
    nodeName: 'p:sp',
  }

  it('produces shared line positions, bullets, and spacing for left-aligned paragraphs', () => {
    const element: TextElement = {
      id: 'text-left',
      index: 0,
      type: 'text',
      slidePart: '/ppt/slides/slide1.xml',
      source,
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
              bullet: { type: 'character', character: '•', color: '#2563eb' },
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

    const layout = layoutTextElement(element)

    expect(layout).not.toBeNull()
    expect(layout?.align).toBe('left')
    expect(layout?.defaultFontSize).toBe(18)
    expect(layout?.lines).toHaveLength(3)
    expect(layout?.lines[0]).toMatchObject({ x: 38, y: 54 })
    expect(layout?.lines[1]).toMatchObject({ x: 74, y: 81 })
    expect(layout?.lines[2]).toMatchObject({ x: 38, y: 116 })
    expect(layout?.lines[0]?.runs[0]).toMatchObject({ text: '• ', style: { color: '#2563eb', fontSize: 1800 } })
    expect(layout?.lines[0]?.runs[1]).toMatchObject({ text: 'Alpha' })
    expect(layout?.lines[1]?.runs[0]).toMatchObject({ text: 'Beta' })
    expect(layout?.lines[2]?.runs[0]).toMatchObject({ text: 'Gamma' })
  })

  it('pushes first-line list content after measured bullet width', () => {
    const element: TextElement = {
      id: 'text-bullet-measured',
      index: 0,
      type: 'text',
      slidePart: '/ppt/slides/slide-bullet-measured.xml',
      source,
      visible: true,
      opacity: 1,
      zIndex: 0,
      transform: { x: 20, y: 10, width: 160, height: 100 },
      text: 'Alpha Beta',
      textBody: {
        paragraphs: [
          {
            text: 'Alpha Beta',
            runs: [{ text: 'Alpha Beta', style: { fontSize: 1600 } }],
            style: {
              bullet: { type: 'character', character: '•' },
              marginLeft: 457200,
              indent: -228600,
              defaultRunStyle: { fontSize: 1600 },
            },
          },
        ],
      },
    }

    const layout = layoutTextElement(element, {
      measureText: ({ text }) => normalizeTextRunText(text).length * 8,
    })

    expect(layout?.lines).toHaveLength(1)
    expect(layout?.lines[0]?.x).toBe(38)
    expect(layout?.lines[0]?.runs[0]).toMatchObject({ text: '• ' })
    expect(layout?.lines[0]?.runs[1]).toMatchObject({ text: 'Alpha Beta' })

    const bulletWidth = normalizeTextRunText(layout?.lines[0]?.runs[0]?.text ?? '').length * 8
    expect((layout?.lines[0]?.x ?? 0) + bulletWidth).toBe(54)
  })

  it('wraps long left-aligned text into multiple measured lines', () => {
    const element: TextElement = {
      id: 'text-wrap',
      index: 0,
      type: 'text',
      slidePart: '/ppt/slides/slide-wrap.xml',
      source,
      visible: true,
      opacity: 1,
      zIndex: 0,
      transform: { x: 10, y: 12, width: 72, height: 120 },
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

    const layout = layoutTextElement(element, {
      measureText: ({ text }) => normalizeTextRunText(text).length * 8,
    })

    expect(layout?.lines).toHaveLength(3)
    expect(layout?.lines[0]?.runs.map((run) => run.text).join('')).toBe('Alpha ')
    expect(layout?.lines[1]?.runs.map((run) => run.text).join('')).toBe('Beta ')
    expect(layout?.lines[2]?.runs.map((run) => run.text).join('')).toBe('Gamma')
    expect(layout?.lines[0]).toMatchObject({ x: 10, y: 28 })
    expect(layout?.lines[1]).toMatchObject({ x: 10, y: 48.8 })
  })

  it('reduces first-line wrap width by measured bullet width', () => {
    const element: TextElement = {
      id: 'text-bullet-wrap',
      index: 0,
      type: 'text',
      slidePart: '/ppt/slides/slide-bullet-wrap.xml',
      source,
      visible: true,
      opacity: 1,
      zIndex: 0,
      transform: { x: 20, y: 10, width: 84, height: 120 },
      text: 'Alpha Beta',
      textBody: {
        paragraphs: [
          {
            text: 'Alpha Beta',
            runs: [{ text: 'Alpha Beta', style: { fontSize: 1600 } }],
            style: {
              bullet: { type: 'character', character: '•' },
              marginLeft: 457200,
              indent: -228600,
              defaultRunStyle: { fontSize: 1600 },
            },
          },
        ],
      },
    }

    const layout = layoutTextElement(element, {
      measureText: ({ text }) => normalizeTextRunText(text).length * 8,
    })

    expect(layout?.lines).toHaveLength(2)
    expect(layout?.lines[0]?.runs.map((run) => run.text).join('')).toBe('• Alpha ')
    expect(layout?.lines[1]?.runs.map((run) => run.text).join('')).toBe('Beta')
    expect(layout?.lines[0]?.x).toBe(38)
    expect(layout?.lines[1]?.x).toBe(74)
  })

  it('splits oversized tokens at character boundaries when needed', () => {
    const element: TextElement = {
      id: 'text-break-word',
      index: 0,
      type: 'text',
      slidePart: '/ppt/slides/slide-break.xml',
      source,
      visible: true,
      opacity: 1,
      zIndex: 0,
      transform: { x: 0, y: 0, width: 36, height: 80 },
      text: 'ABCDEFG',
      textBody: {
        paragraphs: [
          {
            text: 'ABCDEFG',
            runs: [{ text: 'ABCDEFG', style: { fontSize: 1600 } }],
            style: { defaultRunStyle: { fontSize: 1600 } },
          },
        ],
      },
    }

    const layout = layoutTextElement(element, {
      measureText: ({ text }) => normalizeTextRunText(text).length * 8,
    })

    expect(layout?.lines).toHaveLength(2)
    expect(layout?.lines[0]?.runs.map((run) => run.text).join('')).toBe('ABCD')
    expect(layout?.lines[1]?.runs.map((run) => run.text).join('')).toBe('EFG')
  })

  it('repositions wrapped center-aligned lines from measured widths', () => {
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

    const layout = layoutTextElement(element, {
      measureText: ({ text }) => normalizeTextRunText(text).length * 8,
    })

    expect(layout?.align).toBe('center')
    expect(layout?.lines).toHaveLength(2)
    expect(layout?.lines[0]?.runs.map((run) => run.text).join('')).toBe('Alpha ')
    expect(layout?.lines[1]?.runs.map((run) => run.text).join('')).toBe('Beta')
    expect(layout?.lines[0]?.x).toBe(32)
    expect(layout?.lines[1]?.x).toBe(40)
  })

  it('repositions wrapped right-aligned lines from measured widths', () => {
    const element: TextElement = {
      id: 'text-right-wrap',
      index: 0,
      type: 'text',
      slidePart: '/ppt/slides/slide-right-wrap.xml',
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
            style: { align: 'r', defaultRunStyle: { fontSize: 1600 } },
          },
        ],
      },
    }

    const layout = layoutTextElement(element, {
      measureText: ({ text }) => normalizeTextRunText(text).length * 8,
    })

    expect(layout?.align).toBe('right')
    expect(layout?.lines).toHaveLength(2)
    expect(layout?.lines[0]?.x).toBe(44)
    expect(layout?.lines[1]?.x).toBe(60)
  })

  it('increments auto-number bullets across consecutive paragraphs', () => {
    const element: TextElement = {
      id: 'text-auto-number',
      index: 0,
      type: 'text',
      slidePart: '/ppt/slides/slide-auto-number.xml',
      source,
      visible: true,
      opacity: 1,
      zIndex: 0,
      transform: { x: 10, y: 10, width: 240, height: 120 },
      text: 'One\nTwo\nThree',
      textBody: {
        paragraphs: [
          {
            text: 'One',
            runs: [{ text: 'One' }],
            style: { bullet: { type: 'auto-number', autoNumberStartAt: 3 }, defaultRunStyle: { fontSize: 1400 } },
          },
          {
            text: 'Two',
            runs: [{ text: 'Two' }],
            style: { bullet: { type: 'auto-number', autoNumberStartAt: 3 }, defaultRunStyle: { fontSize: 1400 } },
          },
          {
            text: 'Three',
            runs: [{ text: 'Three' }],
            style: { bullet: { type: 'auto-number', autoNumberStartAt: 3 }, defaultRunStyle: { fontSize: 1400 } },
          },
        ],
      },
    }

    const layout = layoutTextElement(element)

    expect(layout?.lines).toHaveLength(3)
    expect(layout?.lines[0]?.runs[0]).toMatchObject({ text: '3. ' })
    expect(layout?.lines[1]?.runs[0]).toMatchObject({ text: '4. ' })
    expect(layout?.lines[2]?.runs[0]).toMatchObject({ text: '5. ' })
  })

  it('restarts auto-numbering after a non-numbered paragraph break', () => {
    const element: TextElement = {
      id: 'text-auto-number-reset',
      index: 0,
      type: 'text',
      slidePart: '/ppt/slides/slide-auto-number-reset.xml',
      source,
      visible: true,
      opacity: 1,
      zIndex: 0,
      transform: { x: 10, y: 10, width: 240, height: 160 },
      text: 'One\nDivider\nTwo',
      textBody: {
        paragraphs: [
          {
            text: 'One',
            runs: [{ text: 'One' }],
            style: { bullet: { type: 'auto-number', autoNumberStartAt: 3 }, defaultRunStyle: { fontSize: 1400 } },
          },
          {
            text: 'Divider',
            runs: [{ text: 'Divider' }],
            style: { bullet: { type: 'none' }, defaultRunStyle: { fontSize: 1400 } },
          },
          {
            text: 'Two',
            runs: [{ text: 'Two' }],
            style: { bullet: { type: 'auto-number', autoNumberStartAt: 7 }, defaultRunStyle: { fontSize: 1400 } },
          },
        ],
      },
    }

    const layout = layoutTextElement(element)

    expect(layout?.lines).toHaveLength(3)
    expect(layout?.lines[0]?.runs[0]).toMatchObject({ text: '3. ' })
    expect(layout?.lines[1]?.runs[0]).toMatchObject({ text: 'Divider' })
    expect(layout?.lines[2]?.runs[0]).toMatchObject({ text: '7. ' })
  })

  it('keeps parent and child auto-number sequences separate across levels', () => {
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

    const layout = layoutTextElement(element)

    expect(layout?.lines).toHaveLength(4)
    expect(layout?.lines[0]?.runs[0]).toMatchObject({ text: '1. ' })
    expect(layout?.lines[1]?.runs[0]).toMatchObject({ text: '1. ' })
    expect(layout?.lines[2]?.runs[0]).toMatchObject({ text: '2. ' })
    expect(layout?.lines[3]?.runs[0]).toMatchObject({ text: '2. ' })
  })

  it('applies text body inset before laying out lines', () => {
    const element: TextElement = {
      id: 'text-inset',
      index: 0,
      type: 'text',
      slidePart: '/ppt/slides/slide-inset.xml',
      source,
      visible: true,
      opacity: 1,
      zIndex: 0,
      transform: { x: 20, y: 30, width: 120, height: 80 },
      text: 'Inset',
      textBody: {
        properties: { inset: { left: 12, top: 8, right: 4, bottom: 6 } },
        paragraphs: [{ text: 'Inset', runs: [{ text: 'Inset', style: { fontSize: 1600 } }], style: { defaultRunStyle: { fontSize: 1600 } } }],
      },
    }

    const layout = layoutTextElement(element)

    expect(layout?.lines[0]).toMatchObject({ x: 32, y: 54 })
  })

  it('keeps large PPT font sizes instead of capping them', () => {
    const element: TextElement = {
      id: 'text-large-font',
      index: 0,
      type: 'text',
      slidePart: '/ppt/slides/slide-large-font.xml',
      source,
      visible: true,
      opacity: 1,
      zIndex: 0,
      transform: { x: 10, y: 20, width: 300, height: 120 },
      text: 'Large title',
      textBody: {
        paragraphs: [{ text: 'Large title', runs: [{ text: 'Large title', style: { fontSize: 4400 } }], style: { defaultRunStyle: { fontSize: 4400 } } }],
      },
    }

    const layout = layoutTextElement(element)

    expect(layout?.defaultFontSize).toBe(44)
    expect(layout?.lines[0]?.y).toBe(64)
    expect(layout?.lines[0]?.runs[0]?.fontSize).toBe(44)
  })

  it('applies vertical text anchoring within the text content box', () => {
    const element: TextElement = {
      id: 'text-vertical-anchor',
      index: 0,
      type: 'text',
      slidePart: '/ppt/slides/slide-vertical-anchor.xml',
      source,
      visible: true,
      opacity: 1,
      zIndex: 0,
      transform: { x: 0, y: 10, width: 200, height: 100 },
      text: 'Middle',
      textBody: {
        properties: { verticalAnchor: 'middle' },
        paragraphs: [{ text: 'Middle', runs: [{ text: 'Middle', style: { fontSize: 2000 } }], style: { defaultRunStyle: { fontSize: 2000 } } }],
      },
    }

    const layout = layoutTextElement(element)

    expect(layout?.lines[0]?.y).toBe(67)
  })

  it('does not auto-wrap when bodyPr disables wrapping', () => {
    const element: TextElement = {
      id: 'text-nowrap',
      index: 0,
      type: 'text',
      slidePart: '/ppt/slides/slide-nowrap.xml',
      source,
      visible: true,
      opacity: 1,
      zIndex: 0,
      transform: { x: 0, y: 0, width: 24, height: 80 },
      text: 'Alpha Beta',
      textBody: {
        properties: { wrap: false },
        paragraphs: [{ text: 'Alpha Beta', runs: [{ text: 'Alpha Beta', style: { fontSize: 1600 } }], style: { defaultRunStyle: { fontSize: 1600 } } }],
      },
    }

    const layout = layoutTextElement(element, {
      measureText: ({ text }) => normalizeTextRunText(text).length * 8,
    })

    expect(layout?.lines).toHaveLength(1)
    expect(layout?.lines[0]?.runs.map((run) => run.text).join('')).toBe('Alpha Beta')
  })

  it('wraps CJK text by character-like segments', () => {
    const element: TextElement = {
      id: 'text-cjk',
      index: 0,
      type: 'text',
      slidePart: '/ppt/slides/slide-cjk.xml',
      source,
      visible: true,
      opacity: 1,
      zIndex: 0,
      transform: { x: 0, y: 0, width: 32, height: 120 },
      text: '企业平台介绍',
      textBody: {
        paragraphs: [{ text: '企业平台介绍', runs: [{ text: '企业平台介绍', style: { fontSize: 1600 } }], style: { defaultRunStyle: { fontSize: 1600 } } }],
      },
    }

    const layout = layoutTextElement(element, {
      measureText: ({ text }) => [...normalizeTextRunText(text)].length * 16,
    })

    expect(layout?.lines.map((line) => line.runs.map((run) => run.text).join(''))).toEqual(['企业', '平台', '介绍'])
  })

  it('keeps CJK line-start punctuation with the previous character', () => {
    const element: TextElement = {
      id: 'text-cjk-punctuation-start',
      index: 0,
      type: 'text',
      slidePart: '/ppt/slides/slide-cjk-punctuation-start.xml',
      source,
      visible: true,
      opacity: 1,
      zIndex: 0,
      transform: { x: 0, y: 0, width: 32, height: 120 },
      text: '标题，正文。',
      textBody: {
        paragraphs: [{ text: '标题，正文。', runs: [{ text: '标题，正文。', style: { fontSize: 1600 } }], style: { defaultRunStyle: { fontSize: 1600 } } }],
      },
    }

    const layout = layoutTextElement(element, {
      measureText: ({ text }) => [...normalizeTextRunText(text)].length * 16,
    })

    expect(layout?.lines.map((line) => line.runs.map((run) => run.text).join(''))).toEqual(['标', '题，', '正', '文。'])
  })

  it('keeps CJK line-end punctuation with the following character', () => {
    const element: TextElement = {
      id: 'text-cjk-punctuation-end',
      index: 0,
      type: 'text',
      slidePart: '/ppt/slides/slide-cjk-punctuation-end.xml',
      source,
      visible: true,
      opacity: 1,
      zIndex: 0,
      transform: { x: 0, y: 0, width: 32, height: 120 },
      text: '标题（正文',
      textBody: {
        paragraphs: [{ text: '标题（正文', runs: [{ text: '标题（正文', style: { fontSize: 1600 } }], style: { defaultRunStyle: { fontSize: 1600 } } }],
      },
    }

    const layout = layoutTextElement(element, {
      measureText: ({ text }) => [...normalizeTextRunText(text)].length * 16,
    })

    expect(layout?.lines.map((line) => line.runs.map((run) => run.text).join(''))).toEqual(['标题', '（正', '文'])
  })

  it('shrinks normal autofit text when content exceeds the box height', () => {
    const element: TextElement = {
      id: 'text-autofit',
      index: 0,
      type: 'text',
      slidePart: '/ppt/slides/slide-autofit.xml',
      source,
      visible: true,
      opacity: 1,
      zIndex: 0,
      transform: { x: 0, y: 0, width: 200, height: 20 },
      text: 'One\nTwo',
      textBody: {
        properties: { autoFit: { type: 'normal', fontScale: 50000 } },
        paragraphs: [
          { text: 'One', runs: [{ text: 'One', style: { fontSize: 2000 } }], style: { defaultRunStyle: { fontSize: 2000 } } },
          { text: 'Two', runs: [{ text: 'Two', style: { fontSize: 2000 } }], style: { defaultRunStyle: { fontSize: 2000 } } },
        ],
      },
    }

    const layout = layoutTextElement(element)

    expect(layout?.defaultFontSize).toBeLessThan(20)
    expect(layout?.lines[0]?.runs[0]?.fontSize).toBeLessThan(20)
  })

  it('uses normal autofit lineSpaceReduction after font scaling', () => {
    const element: TextElement = {
      id: 'text-autofit-line-spacing',
      index: 0,
      type: 'text',
      slidePart: '/ppt/slides/slide-autofit-line-spacing.xml',
      source,
      visible: true,
      opacity: 1,
      zIndex: 0,
      transform: { x: 0, y: 0, width: 200, height: 56 },
      text: 'One\nTwo\nThree',
      textBody: {
        properties: { autoFit: { type: 'normal', fontScale: 100000, lineSpaceReduction: 30000 } },
        paragraphs: [
          { text: 'One', runs: [{ text: 'One', style: { fontSize: 2000 } }], style: { defaultRunStyle: { fontSize: 2000 } } },
          { text: 'Two', runs: [{ text: 'Two', style: { fontSize: 2000 } }], style: { defaultRunStyle: { fontSize: 2000 } } },
          { text: 'Three', runs: [{ text: 'Three', style: { fontSize: 2000 } }], style: { defaultRunStyle: { fontSize: 2000 } } },
        ],
      },
    }

    const layout = layoutTextElement(element)
    const firstGap = (layout?.lines[1]?.y ?? 0) - (layout?.lines[0]?.y ?? 0)

    expect(layout?.defaultFontSize).toBe(20)
    expect(firstGap).toBeCloseTo(18.67)
  })

  it('caps normal autofit lineSpaceReduction at a safe minimum', () => {
    const element: TextElement = {
      id: 'text-autofit-line-spacing-cap',
      index: 0,
      type: 'text',
      slidePart: '/ppt/slides/slide-autofit-line-spacing-cap.xml',
      source,
      visible: true,
      opacity: 1,
      zIndex: 0,
      transform: { x: 0, y: 0, width: 200, height: 20 },
      text: 'One\nTwo',
      textBody: {
        properties: { autoFit: { type: 'normal', fontScale: 100000, lineSpaceReduction: 90000 } },
        paragraphs: [
          { text: 'One', runs: [{ text: 'One', style: { fontSize: 2000 } }], style: { defaultRunStyle: { fontSize: 2000 } } },
          { text: 'Two', runs: [{ text: 'Two', style: { fontSize: 2000 } }], style: { defaultRunStyle: { fontSize: 2000 } } },
        ],
      },
    }

    const layout = layoutTextElement(element)
    const firstGap = (layout?.lines[1]?.y ?? 0) - (layout?.lines[0]?.y ?? 0)

    expect(firstGap).toBeCloseTo(15.6)
  })

  it('treats center-aligned paragraphs as a centered text block', () => {
    const element: TextElement = {
      id: 'text-center',
      index: 0,
      type: 'text',
      slidePart: '/ppt/slides/slide2.xml',
      source,
      visible: true,
      opacity: 1,
      zIndex: 0,
      transform: { x: 24, y: 40, width: 260, height: 120 },
      text: 'Gamma',
      textBody: {
        paragraphs: [
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

    const layout = layoutTextElement(element)

    expect(layout).not.toBeNull()
    expect(layout?.align).toBe('center')
    expect(layout?.lines).toHaveLength(1)
    expect(layout?.lines[0]).toMatchObject({ x: 154, y: 56 })
    expect(layout?.lines[0]?.runs[0]).toMatchObject({ text: 'Gamma' })
  })
})

describe('text helpers', () => {
  it('normalizes tabs and resolves effective font size', () => {
    expect(normalizeTextRunText('A\tB')).toBe('A    B')
    expect(textStyleFontSize({ fontSize: 1600 }, 18)).toBe(16)
    expect(textStyleFontSize(undefined, 18)).toBe(18)
  })
})
