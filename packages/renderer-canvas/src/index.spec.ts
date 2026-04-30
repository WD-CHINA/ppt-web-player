import { describe, expect, it } from 'vitest'

import type { ElementSource, Slide, TextElement } from '@pptx-player/core'

import { renderSlideToCanvas } from './index'

describe('renderSlideToCanvas', () => {
  const source: ElementSource = {
    part: '/ppt/slides/slide1.xml',
    nodeName: 'p:sp',
  }

  it('draws image backgrounds before slide elements', () => {
    const operations: string[] = []
    const context = createMockContext(operations)
    const bitmap = { width: 200, height: 100 } as CanvasImageSource
    const slide: Slide = {
      id: 'slide-image-background',
      index: 0,
      part: '/ppt/slides/slide-image-background.xml',
      relationshipId: 'rIdImageBackground',
      background: {
        type: 'image',
        fill: {
          relationshipId: 'rIdBg',
          imagePart: 'ppt/media/background.png',
          crop: { left: 0.1, top: 0.2, right: 0.3, bottom: 0.1 },
          opacity: 0.5,
          isExternal: false,
        },
      },
      elements: [
        {
          id: 'shape-1',
          index: 0,
          type: 'shape',
          slidePart: '/ppt/slides/slide-image-background.xml',
          source,
          visible: true,
          opacity: 1,
          zIndex: 0,
          transform: { x: 10, y: 20, width: 120, height: 80 },
          fill: { type: 'solid', color: '#22c55e' },
        },
      ],
      diagnostics: [],
    }

    renderSlideToCanvas({
      presentation: { width: 960, height: 540 },
      slide,
      context,
      mediaBitmaps: { 'ppt/media/background.png': bitmap },
      clear: false,
    })

    const backgroundIndex = operations.findIndex((operation) => operation.startsWith('drawImage:20,20,120.'))
    const elementIndex = operations.indexOf('fill:#22c55e')

    expect(operations).toContain('globalAlpha:0.5')
    expect(backgroundIndex).toBeGreaterThan(-1)
    expect(elementIndex).toBeGreaterThan(backgroundIndex)
  })

  it('draws cropped images with source and destination rectangles', () => {
    const operations: string[] = []
    const context = createMockContext(operations)
    const bitmap = { width: 200, height: 100 } as CanvasImageSource
    const slide: Slide = {
      id: 'slide-image-crop',
      index: 0,
      part: '/ppt/slides/slide-image-crop.xml',
      relationshipId: 'rIdImageCrop',
      background: undefined,
      elements: [
        {
          id: 'image-crop',
          index: 0,
          type: 'image',
          slidePart: '/ppt/slides/slide-image-crop.xml',
          source,
          visible: true,
          opacity: 1,
          zIndex: 0,
          transform: { x: 10, y: 20, width: 120, height: 80 },
          relationshipId: 'rIdImage',
          imagePart: 'ppt/media/image1.png',
          crop: { left: 0.1, top: 0.2, right: 0.3, bottom: 0.1 },
          image: { part: 'ppt/media/image1.png', isExternal: false },
          isExternal: false,
        },
      ],
      diagnostics: [],
    }

    renderSlideToCanvas({
      presentation: { width: 960, height: 540 },
      slide,
      context,
      mediaBitmaps: { 'ppt/media/image1.png': bitmap },
      clear: false,
    })

    expect(operations.some((operation) => operation.startsWith('drawImage:20,20,120.'))).toBe(true)
  })

  it('wraps long text into multiple canvas lines', () => {
    const operations: string[] = []
    const context = createMockContext(operations)

    const slide: Slide = {
      id: 'slide-wrap',
      index: 0,
      part: '/ppt/slides/slide-wrap.xml',
      relationshipId: 'rIdWrap',
      background: undefined,
      elements: [
        {
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
        },
      ],
      diagnostics: [],
    }

    renderSlideToCanvas({
      presentation: { width: 960, height: 540 },
      slide,
      context,
      clear: false,
    })

    expect(operations).toContain('fillText:Alpha Beta @12,36')
    expect(operations).toContain('fillText:Gamma@12,56.8')
  })

  it('uses shared left-edge positions for center-aligned wrapped canvas text', () => {
    const operations: string[] = []
    const context = createMockContext(operations)

    const slide: Slide = {
      id: 'slide-center-wrap',
      index: 0,
      part: '/ppt/slides/slide-center-wrap.xml',
      relationshipId: 'rIdCenterWrap',
      background: undefined,
      elements: [
        {
          id: 'text-center-wrap',
          index: 0,
          type: 'text',
          slidePart: '/ppt/slides/slide-center-wrap.xml',
          source,
          visible: true,
          opacity: 1,
          zIndex: 0,
          transform: { x: 20, y: 30, width: 48, height: 100 },
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
        },
      ],
      diagnostics: [],
    }

    renderSlideToCanvas({
      presentation: { width: 960, height: 540 },
      slide,
      context,
      clear: false,
    })

    expect(operations).toContain('fillText:Alpha @26,46')
    expect(operations).toContain('fillText:Beta@32,66.8')
  })

  it('increments auto-number bullets across consecutive canvas paragraphs', () => {
    const operations: string[] = []
    const context = createMockContext(operations)

    const slide: Slide = {
      id: 'slide-auto-number',
      index: 0,
      part: '/ppt/slides/slide-auto-number.xml',
      relationshipId: 'rIdAutoNumber',
      background: undefined,
      elements: [
        {
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
        },
      ],
      diagnostics: [],
    }

    renderSlideToCanvas({
      presentation: { width: 960, height: 540 },
      slide,
      context,
      clear: false,
    })

    expect(operations).toContain('fillText:3. @10,24')
    expect(operations).toContain('fillText:4. @10,42.2')
    expect(operations).toContain('fillText:5. @10,60.4')
  })

  it('keeps parent and child auto-number sequences separate on canvas', () => {
    const operations: string[] = []
    const context = createMockContext(operations)

    const slide: Slide = {
      id: 'slide-auto-number-levels',
      index: 0,
      part: '/ppt/slides/slide-auto-number-levels.xml',
      relationshipId: 'rIdAutoNumberLevels',
      background: undefined,
      elements: [
        {
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
        },
      ],
      diagnostics: [],
    }

    renderSlideToCanvas({ presentation: { width: 960, height: 540 }, slide, context, clear: false })

    expect(operations).toContain('fillText:1. @10,24')
    expect(operations).toContain('fillText:1. @10,42.2')
    expect(operations).toContain('fillText:2. @10,60.4')
    expect(operations).toContain('fillText:2. @10,78.6')
  })

  it('skips hidden and transparent canvas elements', () => {
    const operations: string[] = []
    const context = createMockContext(operations)
    const slide: Slide = {
      id: 'slide-visibility',
      index: 0,
      part: '/ppt/slides/slide-visibility.xml',
      relationshipId: 'rIdVisibility',
      background: undefined,
      elements: [
        {
          id: 'hidden-text',
          index: 0,
          type: 'text',
          slidePart: '/ppt/slides/slide-visibility.xml',
          source,
          visible: false,
          opacity: 1,
          zIndex: 0,
          transform: { x: 10, y: 10, width: 200, height: 60 },
          text: 'Hidden text',
          textBody: { paragraphs: [{ text: 'Hidden text', runs: [{ text: 'Hidden text' }] }] },
        },
        {
          id: 'transparent-text',
          index: 1,
          type: 'text',
          slidePart: '/ppt/slides/slide-visibility.xml',
          source,
          visible: true,
          opacity: 0,
          zIndex: 0,
          transform: { x: 10, y: 80, width: 200, height: 60 },
          text: 'Transparent text',
          textBody: { paragraphs: [{ text: 'Transparent text', runs: [{ text: 'Transparent text' }] }] },
        },
      ],
      diagnostics: [],
    }

    renderSlideToCanvas({ presentation: { width: 960, height: 540 }, slide, context, clear: false })

    expect(operations.some((entry) => entry.includes('Hidden text'))).toBe(false)
    expect(operations.some((entry) => entry.includes('Transparent text'))).toBe(false)
  })

  it('applies element opacity to canvas drawing', () => {
    const operations: string[] = []
    const context = createMockContext(operations)
    const slide: Slide = {
      id: 'slide-opacity',
      index: 0,
      part: '/ppt/slides/slide-opacity.xml',
      relationshipId: 'rIdOpacity',
      background: undefined,
      elements: [
        {
          id: 'opacity-text',
          index: 0,
          type: 'text',
          slidePart: '/ppt/slides/slide-opacity.xml',
          source,
          visible: true,
          opacity: 0.4,
          zIndex: 0,
          transform: { x: 10, y: 10, width: 200, height: 60 },
          text: 'Opacity text',
          textBody: { paragraphs: [{ text: 'Opacity text', runs: [{ text: 'Opacity text' }] }] },
        },
      ],
      diagnostics: [],
    }

    renderSlideToCanvas({ presentation: { width: 960, height: 540 }, slide, context, clear: false })

    expect(operations).toContain('globalAlpha:0.4')
    expect(operations.some((entry) => entry.startsWith('fillText:Opacity text@'))).toBe(true)
  })

  it('applies canvas connector dash and stroke opacity', () => {
    const operations: string[] = []
    const context = createMockContext(operations)
    const slide: Slide = {
      id: 'slide-connector-style',
      index: 0,
      part: '/ppt/slides/slide-connector-style.xml',
      relationshipId: 'rIdConnectorStyle',
      background: undefined,
      elements: [
        {
          id: 'styled-connector',
          index: 0,
          type: 'connector',
          slidePart: '/ppt/slides/slide-connector-style.xml',
          source,
          visible: true,
          opacity: 1,
          zIndex: 0,
          transform: { x: 20, y: 30, width: 120, height: 20 },
          line: { color: '#0f172a', width: 3, opacity: 0.25, dash: 'dash' },
          diagnostics: [],
        },
      ],
      diagnostics: [],
    }

    renderSlideToCanvas({ presentation: { width: 960, height: 540 }, slide, context, clear: false })

    expect(operations).toContain('setLineDash:9,6')
    expect(operations).toContain('globalAlpha:0.25')
    expect(operations).toContain('stroke:#0f172a')
  })

  it('applies canvas shape fill opacity and stroke opacity', () => {
    const operations: string[] = []
    const context = createMockContext(operations)
    const slide: Slide = {
      id: 'slide-shape-opacity',
      index: 0,
      part: '/ppt/slides/slide-shape-opacity.xml',
      relationshipId: 'rIdShapeOpacity',
      background: undefined,
      elements: [
        {
          id: 'opacity-shape',
          index: 0,
          type: 'shape',
          slidePart: '/ppt/slides/slide-shape-opacity.xml',
          source,
          visible: true,
          opacity: 1,
          zIndex: 0,
          transform: { x: 10, y: 10, width: 80, height: 40 },
          fill: { type: 'solid', color: '#22c55e', opacity: 0.5 },
          line: { color: '#166534', width: 2, opacity: 0.75 },
          diagnostics: [],
        },
      ],
      diagnostics: [],
    }

    renderSlideToCanvas({ presentation: { width: 960, height: 540 }, slide, context, clear: false })

    expect(operations).toContain('globalAlpha:0.5')
    expect(operations).toContain('fill:#22c55e')
    expect(operations).toContain('globalAlpha:0.75')
    expect(operations).toContain('stroke:#166534')
  })

  it('renders slide background, shape, connector, image placeholder, and text runs', () => {
    const operations: string[] = []
    const context = createMockContext(operations)

    const textElement: TextElement = {
      id: 'text-1',
      index: 0,
      type: 'text',
      slidePart: '/ppt/slides/slide1.xml',
      source,
      visible: true,
      opacity: 1,
      zIndex: 1,
      transform: { x: 20, y: 30, width: 240, height: 100 },
      text: 'Alpha\nBeta',
      textBody: {
        paragraphs: [
          {
            text: 'Alpha\nBeta',
            runs: [
              { text: 'Alpha', style: { bold: true, fontSize: 1800, color: '#ef4444' } },
              { text: '\nBeta', style: { italic: true, fontSize: 1600 } },
            ],
            style: {
              bullet: { type: 'character', character: '•', color: '#2563eb' },
              marginLeft: 457200,
              indent: -228600,
              defaultRunStyle: { fontSize: 1800 },
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
      background: { type: 'fill', fill: { type: 'solid', color: '#ffffff', opacity: 1 } },
      elements: [
        {
          id: 'shape-1',
          index: 0,
          type: 'shape',
          slidePart: '/ppt/slides/slide1.xml',
          source,
          visible: true,
          opacity: 1,
          zIndex: 0,
          transform: { x: 10, y: 10, width: 80, height: 40 },
          fill: { type: 'solid', color: '#22c55e', opacity: 1 },
          line: { color: '#166534', width: 2 },
          diagnostics: [],
        },
        {
          id: 'connector-1',
          index: 1,
          type: 'connector',
          slidePart: '/ppt/slides/slide1.xml',
          source,
          visible: true,
          opacity: 1,
          zIndex: 0,
          transform: { x: 100, y: 20, width: 90, height: 20 },
          line: { color: '#0f172a', width: 3, dash: 'dash' },
          diagnostics: [],
        },
        {
          id: 'image-1',
          index: 2,
          type: 'image',
          slidePart: '/ppt/slides/slide1.xml',
          source,
          visible: true,
          opacity: 1,
          zIndex: 0,
          transform: { x: 200, y: 20, width: 100, height: 60 },
          relationshipId: 'rIdImg1',
          imagePart: '/ppt/media/image1.png',
          isExternal: false,
          diagnostics: [],
        },
        textElement,
      ],
      diagnostics: [],
    }

    renderSlideToCanvas({
      presentation: { width: 960, height: 540 },
      slide,
      context,
      clear: true,
    })

    expect(operations).toContain('clearRect:0,0,960,540')
    expect(operations).toContain('fillRect:0,0,960,540')
    expect(operations.some((entry) => entry.startsWith('strokeRect:'))).toBe(false)
    expect(operations).toContain('fillText:• @38,48')
    expect(operations).toContain('fillText:Alpha@50,48')
    expect(operations).toContain('fillText:Beta@74,71.4')
    expect(operations).toContain('fillText:/ppt/media/image1.png@208,28')
  })
})

function createMockContext(operations: string[]): CanvasRenderingContext2D {
  const state = {
    fillStyle: '#000000',
    strokeStyle: '#000000',
    lineWidth: 1,
    font: '10px sans-serif',
    textAlign: 'left' as CanvasTextAlign,
    textBaseline: 'alphabetic' as CanvasTextBaseline,
    globalAlpha: 1,
  }
  const stack: Array<typeof state> = []

  return {
    get fillStyle() {
      return state.fillStyle
    },
    set fillStyle(value) {
      state.fillStyle = String(value)
    },
    get strokeStyle() {
      return state.strokeStyle
    },
    set strokeStyle(value) {
      state.strokeStyle = String(value)
    },
    get lineWidth() {
      return state.lineWidth
    },
    set lineWidth(value) {
      state.lineWidth = value
    },
    get font() {
      return state.font
    },
    set font(value) {
      state.font = value
    },
    get textAlign() {
      return state.textAlign
    },
    set textAlign(value) {
      state.textAlign = value
    },
    get textBaseline() {
      return state.textBaseline
    },
    set textBaseline(value) {
      state.textBaseline = value
    },
    get globalAlpha() {
      return state.globalAlpha
    },
    set globalAlpha(value) {
      state.globalAlpha = value
      operations.push(`globalAlpha:${value}`)
    },
    save() {
      stack.push({ ...state })
      operations.push('save')
    },
    restore() {
      const previous = stack.pop()

      if (previous) {
        Object.assign(state, previous)
      }

      operations.push('restore')
    },
    clearRect(x: number, y: number, width: number, height: number) {
      operations.push(`clearRect:${x},${y},${width},${height}`)
    },
    fillRect(x: number, y: number, width: number, height: number) {
      operations.push(`fillRect:${x},${y},${width},${height}`)
    },
    strokeRect(x: number, y: number, width: number, height: number) {
      operations.push(`strokeRect:${x},${y},${width},${height}`)
    },
    beginPath() {
      operations.push('beginPath')
    },
    moveTo() {},
    lineTo() {},
    ellipse() {},
    arcTo() {},
    closePath() {},
    fill() {
      operations.push(`fill:${state.fillStyle}`)
    },
    stroke() {
      operations.push(`stroke:${state.strokeStyle}`)
    },
    setLineDash(segments: number[]) {
      operations.push(`setLineDash:${segments.join(',')}`)
    },
    drawImage(...args: unknown[]) {
      operations.push(`drawImage:${args.slice(1).join(',')}`)
    },
    fillText(text: string, x: number, y: number) {
      operations.push(`fillText:${String(text)}@${x},${y}`)
    },
    measureText(text: string) {
      return { width: String(text).length * 6 } as TextMetrics
    },
  } as unknown as CanvasRenderingContext2D
}
