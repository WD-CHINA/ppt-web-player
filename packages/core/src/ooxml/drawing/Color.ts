import type { PresentationTheme, ThemeColorScheme } from '../../model/Presentation'
import type { XmlNode } from '../../xml/XmlNode'
import * as xml from '../../xml/XmlQuery'

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

export function parseDrawingColor(node: XmlNode | null | undefined, theme?: PresentationTheme): string | undefined {
  const srgbColor = normalizeSrgbColor(xml.attr(xml.child(node, 'a:srgbClr'), 'val'))

  if (srgbColor) {
    return srgbColor
  }

  return resolveThemeColor(xml.attr(xml.child(node, 'a:schemeClr'), 'val'), theme?.colorScheme)
}

export function normalizeSrgbColor(value: string | undefined): string | undefined {
  if (!value || !/^[0-9a-f]{6}$/i.test(value)) {
    return undefined
  }

  return `#${value.toUpperCase()}`
}

export function normalizeAlpha(value: string | undefined): number | undefined {
  if (!value) {
    return undefined
  }

  const alpha = Number(value)

  if (!Number.isFinite(alpha)) {
    return undefined
  }

  return Math.min(Math.max(alpha, 0), 100000) / 100000
}

export function resolveThemeColor(value: string | undefined, colorScheme?: ThemeColorScheme): string | undefined {
  if (!value || !colorScheme) {
    return undefined
  }

  const mappedKey = SCHEME_COLOR_MAP[value as keyof typeof SCHEME_COLOR_MAP]
  return mappedKey ? colorScheme[mappedKey] : undefined
}
