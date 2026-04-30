import type { PresentationTheme, ThemeColorScheme, ThemeFontScheme } from '../../model/Presentation'
import type { PptxPackage } from '../../package/PptxPackage'
import type { XmlNode } from '../../xml/XmlNode'
import * as xml from '../../xml/XmlQuery'
import { normalizeSrgbColor } from '../drawing/Fill'

export async function parseTheme(pptx: PptxPackage, themePart: string): Promise<PresentationTheme> {
  const root = await pptx.getXml(themePart)

  return {
    part: themePart,
    name: xml.attr(root, 'name'),
    colorScheme: root ? parseColorScheme(root) : undefined,
    fontScheme: root ? parseFontScheme(root) : undefined,
  }
}

function parseColorScheme(root: XmlNode): ThemeColorScheme | undefined {
  const colorScheme = xml.path(root, ['a:themeElements', 'a:clrScheme'])

  if (!colorScheme) {
    return undefined
  }

  return {
    dark1: parseSchemeColor(colorScheme, 'a:dk1'),
    light1: parseSchemeColor(colorScheme, 'a:lt1'),
    dark2: parseSchemeColor(colorScheme, 'a:dk2'),
    light2: parseSchemeColor(colorScheme, 'a:lt2'),
    accent1: parseSchemeColor(colorScheme, 'a:accent1'),
    accent2: parseSchemeColor(colorScheme, 'a:accent2'),
    accent3: parseSchemeColor(colorScheme, 'a:accent3'),
    accent4: parseSchemeColor(colorScheme, 'a:accent4'),
    accent5: parseSchemeColor(colorScheme, 'a:accent5'),
    accent6: parseSchemeColor(colorScheme, 'a:accent6'),
    hyperlink: parseSchemeColor(colorScheme, 'a:hlink'),
    followedHyperlink: parseSchemeColor(colorScheme, 'a:folHlink'),
  }
}

function parseFontScheme(root: XmlNode): ThemeFontScheme | undefined {
  const fontScheme = xml.path(root, ['a:themeElements', 'a:fontScheme'])

  if (!fontScheme) {
    return undefined
  }

  return {
    majorLatin: xml.attr(xml.path(fontScheme, ['a:majorFont', 'a:latin']), 'typeface'),
    minorLatin: xml.attr(xml.path(fontScheme, ['a:minorFont', 'a:latin']), 'typeface'),
  }
}

function parseSchemeColor(node: XmlNode, childName: string): string | undefined {
  const schemeNode = xml.child(node, childName)
  const srgb = normalizeSrgbColor(xml.attr(xml.child(schemeNode, 'a:srgbClr'), 'val'))

  if (srgb) {
    return srgb
  }

  return normalizeSrgbColor(xml.attr(xml.child(schemeNode, 'a:sysClr'), 'lastClr'))
}
