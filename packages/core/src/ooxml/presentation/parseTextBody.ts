import type {
  BulletStyle,
  Paragraph,
  ParagraphStyle,
  PresentationTheme,
  TextBody,
  TextBodyProperties,
  TextRun,
  TextSpacing,
  TextStyle,
  TextStyleDefaults,
} from '../../model/Presentation'
import type { XmlNode } from '../../xml/XmlNode'
import * as xml from '../../xml/XmlQuery'
import { parseDrawingColor } from '../drawing/Color'
import { emuToPx } from '../drawing/Transform'

export function parseTextBody(node: XmlNode, theme?: PresentationTheme, inheritedDefaults?: TextStyleDefaults): TextBody | undefined {
  const textBody = xml.child(node, 'p:txBody')

  if (!textBody) {
    return undefined
  }

  const paragraphs = xml
    .children(textBody, 'a:p')
    .map((paragraph) => parseParagraph(paragraph, theme, inheritedDefaults))
    .filter((paragraph) => paragraph.text.length > 0)

  if (paragraphs.length === 0) {
    return undefined
  }

  const properties = parseTextBodyProperties(xml.child(textBody, 'a:bodyPr'))

  return { paragraphs, ...(properties ? { properties } : {}) }
}

export function parseTextBodyDefaults(node: XmlNode, theme?: PresentationTheme): TextStyleDefaults | undefined {
  const textBody = xml.child(node, 'p:txBody')

  if (!textBody) {
    return undefined
  }

  const paragraphs: Record<number, ParagraphStyle> = {}
  const listStyle = xml.child(textBody, 'a:lstStyle')

  for (let index = 1; index <= 9; index += 1) {
    const style = parseParagraphStyle(xml.child(listStyle, `a:lvl${index}pPr`), theme)

    if (style) {
      paragraphs[index - 1] = style
    }
  }

  const firstParagraphStyle = parseParagraphStyle(xml.child(xml.child(textBody, 'a:p'), 'a:pPr'), theme)

  if (firstParagraphStyle) {
    const mergedStyle = mergeParagraphStyle(paragraphs[firstParagraphStyle.level ?? 0], firstParagraphStyle)

    if (mergedStyle) {
      paragraphs[firstParagraphStyle.level ?? 0] = mergedStyle
    }
  }

  return Object.keys(paragraphs).length > 0 ? { paragraphs } : undefined
}

function parseTextBodyProperties(node: XmlNode | null): TextBodyProperties | undefined {
  if (!node) {
    return undefined
  }

  const inset = parseTextInsets(node)
  const wrap = xml.attr(node, 'wrap') === 'none' ? false : undefined
  const verticalAnchor = parseVerticalAnchor(xml.attr(node, 'anchor'))
  const autoFit = parseTextAutoFit(node)
  const properties: TextBodyProperties = {
    ...(inset ? { inset } : {}),
    ...(wrap !== undefined ? { wrap } : {}),
    ...(verticalAnchor ? { verticalAnchor } : {}),
    ...(autoFit ? { autoFit } : {}),
  }

  return Object.keys(properties).length > 0 ? properties : undefined
}

function parseTextInsets(node: XmlNode): TextBodyProperties['inset'] | undefined {
  const left = parseNumber(xml.attr(node, 'lIns'))
  const top = parseNumber(xml.attr(node, 'tIns'))
  const right = parseNumber(xml.attr(node, 'rIns'))
  const bottom = parseNumber(xml.attr(node, 'bIns'))

  if (left === undefined && top === undefined && right === undefined && bottom === undefined) {
    return undefined
  }

  return {
    left: left !== undefined ? emuToPx(left) : 0,
    top: top !== undefined ? emuToPx(top) : 0,
    right: right !== undefined ? emuToPx(right) : 0,
    bottom: bottom !== undefined ? emuToPx(bottom) : 0,
  }
}

function parseVerticalAnchor(value: string | undefined): TextBodyProperties['verticalAnchor'] | undefined {
  if (value === 'ctr') {
    return 'middle'
  }

  if (value === 'b') {
    return 'bottom'
  }

  if (value === 't') {
    return 'top'
  }

  return undefined
}

function parseTextAutoFit(node: XmlNode): TextBodyProperties['autoFit'] | undefined {
  if (xml.child(node, 'a:noAutofit')) {
    return { type: 'none' }
  }

  if (xml.child(node, 'a:spAutoFit')) {
    return { type: 'shape' }
  }

  const normalAutoFit = xml.child(node, 'a:normAutofit')

  if (!normalAutoFit) {
    return undefined
  }

  const fontScale = parseNumber(xml.attr(normalAutoFit, 'fontScale'))
  const lineSpaceReduction = parseNumber(xml.attr(normalAutoFit, 'lnSpcReduction'))

  return {
    type: 'normal',
    ...(fontScale !== undefined ? { fontScale } : {}),
    ...(lineSpaceReduction !== undefined ? { lineSpaceReduction } : {}),
  }
}

function parseParagraph(node: XmlNode, theme: PresentationTheme | undefined, inheritedDefaults?: TextStyleDefaults): Paragraph {
  const directStyle = parseParagraphStyle(xml.child(node, 'a:pPr'), theme)
  const level = directStyle?.level ?? 0
  const inheritedStyle = inheritedDefaults?.paragraphs[level]
  const style = mergeParagraphStyle(inheritedStyle, directStyle)
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

function parseParagraphStyle(properties: XmlNode | null, theme?: PresentationTheme): ParagraphStyle | undefined {
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
  const bulletColor = parseBulletColor(node, theme)
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

function parseBulletColor(node: XmlNode, theme?: PresentationTheme): string | undefined {
  return parseDrawingColor(xml.child(node, 'a:buClr'), theme)
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
  const color = parseTextColor(node, theme)

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

function parseTextColor(node: XmlNode, theme?: PresentationTheme): string | undefined {
  return parseDrawingColor(xml.child(node, 'a:solidFill'), theme)
}

function mergeParagraphStyle(base: ParagraphStyle | undefined, override: ParagraphStyle | undefined): ParagraphStyle | undefined {
  if (!base && !override) {
    return undefined
  }

  return {
    ...base,
    ...override,
    defaultRunStyle: mergeTextStyle(base?.defaultRunStyle, override?.defaultRunStyle),
  }
}

function mergeTextStyle(base: TextStyle | undefined, override: TextStyle | undefined): TextStyle | undefined {
  if (!base && !override) {
    return undefined
  }

  return {
    ...base,
    ...override,
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
