import { parseFill } from '../drawing/Fill'
import type { Diagnostic } from '../../diagnostics/Diagnostic'
import type { PresentationTheme, SlideBackground } from '../../model/Presentation'
import type { PptxPackage } from '../../package/PptxPackage'
import type { XmlNode } from '../../xml/XmlNode'
import * as xml from '../../xml/XmlQuery'
import { parseImageFill } from './parseImageFill'

export interface ParseSlideBackgroundOptions {
  pptx: PptxPackage
  part: string
  root: XmlNode
  diagnostics: Diagnostic[]
  theme?: PresentationTheme
  slideIndex?: number
  inheritedBackground?: SlideBackground
}

export async function parseSlideBackground(options: ParseSlideBackgroundOptions): Promise<SlideBackground | undefined> {
  const background = xml.path(options.root, ['p:cSld', 'p:bg'])

  if (!background) {
    return options.inheritedBackground
  }

  const backgroundProperties = xml.child(background, 'p:bgPr')
  const imageFill = backgroundProperties
    ? await parseImageFill(options.pptx, options.part, backgroundProperties, options.diagnostics, {
        slideIndex: options.slideIndex,
        missingMessage: '图片背景缺少 r:embed。',
        externalMessage: '外部图片背景暂不加载。',
      })
    : undefined

  if (imageFill) {
    const { diagnostics: _diagnostics, ...fill } = imageFill
    return { type: 'image', fill }
  }

  const fill = parseFill(background, options.theme)

  return fill ? { type: 'fill', fill } : options.inheritedBackground
}
