import { afterEach, describe, expect, it, vi } from 'vitest'
import { mount } from '@vue/test-utils'

import App from '../App.vue'

const noopMock = () => vi.fn<() => void>()

afterEach(() => {
  vi.restoreAllMocks()
  vi.unstubAllGlobals()
})

vi.mock('@pptx-player/core', async (importOriginal) => {
  const actual = await importOriginal<typeof import('@pptx-player/core')>()

  return {
    ...actual,
    parsePptx: vi.fn<typeof actual.parsePptx>(async () => ({
      presentation: {
        id: 'test-presentation',
        width: 960,
        height: 540,
        slideMasters: [],
        slideLayouts: [],
        slides: [
          {
            id: 'slide-1',
            index: 0,
            part: 'ppt/slides/slide1.xml',
            relationshipId: 'rId1',
            elements: [],
            diagnostics: [],
          },
        ],
        diagnostics: [],
        metadata: {
          presentationPart: 'ppt/presentation.xml',
          slideSize: { widthEmu: 9144000, heightEmu: 5143500 },
        },
      },
      diagnostics: [
        {
          code: 'STYLE_INHERITANCE_INCOMPLETE',
          severity: 'info',
          message: '未找到匹配的 layout/master placeholder 样式，已仅使用 slide 直接样式。',
          part: 'ppt/slides/slide1.xml',
          slideIndex: 0,
          elementId: 'element-4',
          detail: {
            placeholderType: 'subtitle',
            placeholderIndex: '9',
          },
        },
        {
          code: 'IMAGE_RELATIONSHIP_NOT_FOUND',
          severity: 'warning',
          message: '缺少图片关系。',
          part: 'ppt/slides/slide1.xml',
          slideIndex: 0,
        },
      ],
      media: {},
    })),
  }
})

describe('App', () => {
  it('renders the Phase 0 parser UI', () => {
    const wrapper = mount(App)

    expect(wrapper.text()).toContain('企业可控 PPTX Web Player')
    expect(wrapper.text()).toContain('上传本地 PPTX')
    expect(wrapper.text()).toContain('Diagnostics')
    expect(wrapper.text()).toContain('区级平台介绍')
    expect(wrapper.text()).toContain('Slidesgo AI Tech')
    expect(wrapper.text()).toContain('请选择一个样本或上传 PPTX 文件开始解析')
  })

  it('shows diagnostic context and filters diagnostics by severity', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => ({
        ok: true,
        arrayBuffer: async () => new ArrayBuffer(0),
      })),
    )
    vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue({
      clearRect: noopMock(),
      save: noopMock(),
      restore: noopMock(),
      scale: noopMock(),
      fillRect: noopMock(),
      strokeRect: noopMock(),
      beginPath: noopMock(),
      rect: noopMock(),
      ellipse: noopMock(),
      moveTo: noopMock(),
      lineTo: noopMock(),
      closePath: noopMock(),
      fill: noopMock(),
      stroke: noopMock(),
      setLineDash: noopMock(),
      drawImage: noopMock(),
      fillText: noopMock(),
      measureText: vi.fn<() => TextMetrics>(() => ({ width: 0 }) as TextMetrics),
    } as unknown as CanvasRenderingContext2D)

    const wrapper = mount(App)

    await wrapper.find('.sample-list button').trigger('click')
    await vi.dynamicImportSettled()

    expect(wrapper.text()).toContain('共 2 条 · info 1 · warning 1 · error 0')
    expect(wrapper.text()).toContain('STYLE_INHERITANCE_INCOMPLETE')
    expect(wrapper.text()).toContain('IMAGE_RELATIONSHIP_NOT_FOUND')
    expect(wrapper.text()).toContain('Slide')
    expect(wrapper.text()).toContain('#1')
    expect(wrapper.text()).toContain('Element')
    expect(wrapper.text()).toContain('element-4')
    expect(wrapper.text()).toContain('placeholderType')

    await wrapper.find('select').setValue('warning')

    const diagnosticItems = wrapper.findAll('.diagnostics-panel > ul > li')
    const warningItem = diagnosticItems[0]!

    expect(diagnosticItems).toHaveLength(1)
    expect(warningItem).toBeDefined()
    expect(warningItem.text()).not.toContain('STYLE_INHERITANCE_INCOMPLETE')
    expect(warningItem.text()).toContain('IMAGE_RELATIONSHIP_NOT_FOUND')
  })
})
