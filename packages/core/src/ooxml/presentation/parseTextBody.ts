import type {
  BulletStyle,
  Paragraph,
  ParagraphStyle,
  PresentationTheme,
  TextBody,
  TextRun,
  TextSpacing,
  TextStyle,
  ThemeColorScheme,
} from '../../model/Presentation'
import type { XmlNode } from '../../xml/XmlNode'
import * as xml from '../../xml/XmlQuery'
import { normalizeSrgbColor } from '../drawing/Fill'

export function parseTextBody(node: XmlNode, theme?: PresentationTheme): TextBody | undefined {
  const textBody = xml.child(node, 'p:txBody')

  if (!textBody) {
    return undefined
  }

  const paragraphs = xml.children(textBody, 'a:p').map((paragraph) => parseParagraph(paragraph, theme)).filter((paragraph) => paragraph.text.length > 0)

  if (paragraphs.length === 0) {
    return undefined
  }

  return { paragraphs }
}

function parseParagraph(node: XmlNode, theme?: PresentationTheme): Paragraph {
  const style = parseParagraphStyle(node, theme)
  const runs = collectRuns(node, theme, style?.defaultRunStyle)

  return {
    runs,
    text: runs.map((run) => run.text).join(''),
    ...(style ? { style } : {}),
  }
}

function collectRuns(node: XmlNode, theme: PresentationTheme | undefined, defaultRunStyle?: TextStyle): TextRun[] {
  const runs: TextRun[] = []

  for (const child of xml.children(node)) {
    if (child.name === 'a:r') {
      const run = parseTextRun(child, theme, defaultRunStyle)

      if (run) {
        runs.push(run)
      }

      continue
    }

    if (child.name === 'a:fld') {
      const run = parseFieldRun(child, theme, defaultRunStyle)

      if (run) {
        runs.push(run)
      }

      continue
    }

    if (child.name === 'a:br') {
      runs.push({ text: '\n' })
      continue
    }

    if (child.name === 'a:tab') {
      runs.push({ text: '\t', ...(defaultRunStyle ? { style: defaultRunStyle } : {}) })
      continue
    }

    if (child.name === 'a:t') {
      const text = normalizeText(xml.text(child))

      if (text) {
        runs.push({ text, ...(defaultRunStyle ? { style: defaultRunStyle } : {}) })
      }
    }
  }

  return runs
}

function parseTextRun(node: XmlNode, theme: PresentationTheme | undefined, defaultRunStyle?: TextStyle): TextRun | null {
  const text = normalizeText(xml.text(xml.child(node, 'a:t')))

  if (!text) {
    return null
  }

  const style = mergeTextStyle(defaultRunStyle, parseTextStyle(xml.child(node, 'a:rPr'), theme))

  return {
    text,
    ...(style ? { style } : {}),
  }
}

function parseFieldRun(node: XmlNode, theme: PresentationTheme | undefined, defaultRunStyle?: TextStyle): TextRun | null {
  const text = normalizeText(xml.text(node))

  if (!text) {
    return null
  }

  const style = mergeTextStyle(defaultRunStyle, parseTextStyle(xml.child(node, 'a:rPr'), theme))

  return {
    text,
    ...(style ? { style } : {}),
  }
}

function parseParagraphStyle(node: XmlNode, theme?: PresentationTheme): ParagraphStyle | undefined {
  const properties = xml.child(node, 'a:pPr')

  if (!properties) {
    return undefined
  }

  const style: ParagraphStyle = {}
  const align = xml.attr(properties, 'algn')
  const level = parseNumber(xml.attr(properties, 'lvl'))
  const indent = parseNumber(xml.attr(properties, 'indent'))
  const marginLeft = parseNumber(xml.attr(properties, 'marL'))
  const marginRight = parseNumber(xml.attr(properties, 'marR'))
  const defaultTabSize = parseNumber(xml.attr(properties, 'defTabSz'))
  const rtl = parseBooleanFlag(xml.attr(properties, 'rtl'))
  const bullet = parseBulletStyle(properties, theme)
  const lineSpacing = parseSpacing(xml.child(properties, 'a:lnSpc'))
  const spaceBefore = parseSpacing(xml.child(properties, 'a:spcBef'))
  const spaceAfter = parseSpacing(xml.child(properties, 'a:spcAft'))
  const defaultRunStyle = parseTextStyle(xml.child(properties, 'a:defRPr'), theme)

  if (align) {
    style.align = align
  }

  if (level !== undefined) {
    style.level = level
  }

  if (indent !== undefined) {
    style.indent = indent
  }

  if (marginLeft !== undefined) {
    style.marginLeft = marginLeft
  }

  if (marginRight !== undefined) {
    style.marginRight = marginRight
  }

  if (defaultTabSize !== undefined) {
    style.defaultTabSize = defaultTabSize
  }

  if (rtl !== undefined) {
    style.rtl = rtl
  }

  if (bullet) {
    style.bullet = bullet
  }

  if (lineSpacing) {
    style.lineSpacing = lineSpacing
  }

  if (spaceBefore) {
    style.spaceBefore = spaceBefore
  }

  if (spaceAfter) {
    style.spaceAfter = spaceAfter
  }

  if (defaultRunStyle) {
    style.defaultRunStyle = defaultRunStyle
  }

  return Object.keys(style).length > 0 ? style : undefined
}

function parseBulletStyle(node: XmlNode, theme?: PresentationTheme): BulletStyle | undefined {
  const bulletColor = parseBulletColor(node, theme?.colorScheme)
  const bulletFontFace = xml.attr(xml.child(node, 'a:buFont'), 'typeface')
  const bulletFontSize = parseBulletSize(node)

  if (xml.child(node, 'a:buNone')) {
    return {
      type: 'none',
      ...(bulletFontFace ? { fontFace: bulletFontFace } : {}),
      ...(bulletColor ? { color: bulletColor } : {}),
      ...(bulletFontSize !== undefined ? { fontSize: bulletFontSize } : {}),
    }
  }

  const bulletChar = xml.attr(xml.child(node, 'a:buChar'), 'char')

  if (bulletChar) {
    return {
      type: 'character',
      character: bulletChar,
      ...(bulletFontFace ? { fontFace: bulletFontFace } : {}),
      ...(bulletColor ? { color: bulletColor } : {}),
      ...(bulletFontSize !== undefined ? { fontSize: bulletFontSize } : {}),
    }
  }

  const autoNumber = xml.child(node, 'a:buAutoNum')

  if (!autoNumber) {
    return undefined
  }

  const autoNumberScheme = xml.attr(autoNumber, 'type')
  const autoNumberStartAt = parseNumber(xml.attr(autoNumber, 'startAt'))

  return {
    type: 'auto-number',
    ...(autoNumberScheme ? { autoNumberScheme } : {}),
    ...(autoNumberStartAt !== undefined ? { autoNumberStartAt } : {}),
    ...(bulletFontFace ? { fontFace: bulletFontFace } : {}),
    ...(bulletColor ? { color: bulletColor } : {}),
    ...(bulletFontSize !== undefined ? { fontSize: bulletFontSize } : {}),
  }
}

function parseBulletColor(node: XmlNode, colorScheme?: ThemeColorScheme): string | undefined {
  const bulletColorNode = xml.child(node, 'a:buClr')

  if (!bulletColorNode) {
    return undefined
  }

  const srgbColor = normalizeSrgbColor(xml.attr(xml.child(bulletColorNode, 'a:srgbClr'), 'val'))

  if (srgbColor) {
    return srgbColor
  }

  const schemeColor = xml.attr(xml.child(bulletColorNode, 'a:schemeClr'), 'val')

  if (!schemeColor || !colorScheme) {
    return undefined
  }

  return resolveThemeColor(schemeColor, colorScheme)
}

function parseBulletSize(node: XmlNode): number | undefined {
  const bulletSizePoints = parseNumber(xml.attr(xml.child(node, 'a:buSzPts'), 'val'))

  if (bulletSizePoints !== undefined) {
    return bulletSizePoints
  }

  const bulletSizePercent = parseNumber(xml.attr(xml.child(node, 'a:buSzPct'), 'val'))

  if (bulletSizePercent !== undefined) {
    return bulletSizePercent
  }

  return xml.child(node, 'a:buSzTx') ? 0 : undefined
}

function parseSpacing(node: XmlNode | null): TextSpacing | undefined {
  if (!node) {
    return undefined
  }

  const points = parseNumber(xml.attr(xml.child(node, 'a:spcPts'), 'val'))
  const percent = parseNumber(xml.attr(xml.child(node, 'a:spcPct'), 'val'))

  if (points === undefined && percent === undefined) {
    return undefined
  }

  return {
    ...(points !== undefined ? { points } : {}),
    ...(percent !== undefined ? { percent } : {}),
  }
}

function parseTextStyle(node: XmlNode | null, theme?: PresentationTheme): TextStyle | undefined {
  if (!node) {
    return undefined
  }

  const style: TextStyle = {}
  const bold = parseBooleanFlag(xml.attr(node, 'b'))
  const italic = parseBooleanFlag(xml.attr(node, 'i'))
  const underline = xml.attr(node, 'u')
  const fontSize = parseNumber(xml.attr(node, 'sz'))
  const fontFace = xml.attr(node, 'latin:typeface') ?? xml.attr(xml.child(node, 'a:latin'), 'typeface')
  const color = parseTextColor(node, theme?.colorScheme)

  if (bold !== undefined) {
    style.bold = bold
  }

  if (italic !== undefined) {
    style.italic = italic
  }

  if (underline && underline !== 'none') {
    style.underline = underline
  }

  if (fontSize !== undefined) {
    style.fontSize = fontSize
  }

  if (fontFace) {
    style.fontFace = fontFace
  }

  if (color) {
    style.color = color
  }

  return Object.keys(style).length > 0 ? style : undefined
}

function parseTextColor(node: XmlNode, colorScheme?: ThemeColorScheme): string | undefined {
  const solidFill = xml.child(node, 'a:solidFill')
  const srgbColor = normalizeSrgbColor(xml.attr(xml.child(solidFill, 'a:srgbClr'), 'val'))

  if (srgbColor) {
    return srgbColor
  }

  const schemeColor = xml.attr(xml.child(solidFill, 'a:schemeClr'), 'val')

  if (!schemeColor || !colorScheme) {
    return undefined
  }

  return resolveThemeColor(schemeColor, colorScheme)
}

function resolveThemeColor(value: string, colorScheme: ThemeColorScheme): string | undefined {
  const mappedKey = SCHEME_COLOR_MAP[value as keyof typeof SCHEME_COLOR_MAP]
  return mappedKey ? colorScheme[mappedKey] : undefined
}

function mergeTextStyle(base: TextStyle | undefined, override: TextStyle | undefined): TextStyle | undefined {
  if (!base && !override) {
    return undefined
  }

  return {
    ...(base ?? {}),
    ...(override ?? {}),
  }
}

function parseBooleanFlag(value: string | undefined): boolean | undefined {
  if (value === undefined) {
    return undefined
  }

  return value === '1' || value === 'true'
}

function parseNumber(value: string | undefined): number | undefined {
  if (value === undefined) {
    return undefined
  }

  const parsed = Number(value)
  return Number.isFinite(parsed) ? parsed : undefined
}

function normalizeText(value: string): string {
  return value.trim()
}

const SCHEME_COLOR_MAP = {
  dk1: 'dark1',
  lt1: 'light1',
  dk2: 'dark2',
  lt2: 'light2',
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
  hlink: 'hyperlink',
  folHlink: 'followedHyperlink',
} as const
