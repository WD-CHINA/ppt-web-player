import type { Diagnostic } from '../diagnostics/Diagnostic'

export interface Presentation {
  id: string
  width: number
  height: number
  theme?: PresentationTheme
  slideMasters: SlideMaster[]
  slideLayouts: SlideLayout[]
  slides: Slide[]
  diagnostics: Diagnostic[]
  metadata: PresentationMetadata
}

export interface PresentationMetadata {
  presentationPart: string
}

export interface PresentationTheme {
  part: string
  name?: string
  colorScheme?: ThemeColorScheme
  fontScheme?: ThemeFontScheme
}

export interface ThemeColorScheme {
  dark1?: string
  light1?: string
  dark2?: string
  light2?: string
  accent1?: string
  accent2?: string
  accent3?: string
  accent4?: string
  accent5?: string
  accent6?: string
  hyperlink?: string
  followedHyperlink?: string
}

export interface ThemeFontScheme {
  majorLatin?: string
  minorLatin?: string
}

export interface SlideMaster {
  id: string
  part: string
  relationshipId: string
  themePart?: string
  layoutParts: string[]
  defaults?: SlideStyleDefaults
}

export interface SlideLayout {
  id: string
  part: string
  relationshipId?: string
  masterPart?: string
  name?: string
  defaults?: SlideStyleDefaults
}

export interface SlideStyleDefaults {
  background?: SlideBackground
  placeholders: PlaceholderStyle[]
}

export interface PlaceholderStyle {
  type?: string
  index?: string
  name?: string
  fill?: Fill
  line?: LineStyle
  text?: TextStyleDefaults
}

export interface TextStyleDefaults {
  paragraphs: Record<number, ParagraphStyle>
}

export interface Slide {
  id: string
  index: number
  part: string
  relationshipId: string
  name?: string
  background?: SlideBackground
  layoutPart?: string
  masterPart?: string
  themePart?: string
  elements: SlideElement[]
  diagnostics: Diagnostic[]
}

export type SlideElement = TextElement | ImageElement | ShapeElement | ConnectorElement | UnknownElement

export type SlideBackground = FillBackground | ImageBackground

export interface FillBackground {
  type: 'fill'
  fill: Fill
}

export interface ImageBackground {
  type: 'image'
  fill: ImageFill
}

export interface ImageFill {
  relationshipId?: string
  imagePart?: string
  crop?: ImageCrop
  opacity?: number
  isExternal: boolean
}

export type Fill = SolidFill | NoFill

export interface SolidFill {
  type: 'solid'
  color: string
  opacity?: number
}

export interface NoFill {
  type: 'none'
}

export interface LineStyle {
  color?: string
  width?: number
  opacity?: number
  dash?: string
  headEnd?: LineEndStyle
  tailEnd?: LineEndStyle
}

export interface LineEndStyle {
  type: string
  width?: string
  length?: string
}

export type ShapeGeometry = PresetShapeGeometry

export interface PresetShapeGeometry {
  type: 'preset'
  preset: string
}

export interface Transform {
  x: number
  y: number
  width: number
  height: number
}

export interface ElementSource {
  part: string
  nodeName: string
}

export interface TextBody {
  paragraphs: Paragraph[]
  properties?: TextBodyProperties
}

export interface TextBodyProperties {
  inset?: TextInsets
  autoFit?: TextAutoFit
  wrap?: boolean
  verticalAnchor?: 'top' | 'middle' | 'bottom'
}

export interface TextInsets {
  left: number
  top: number
  right: number
  bottom: number
}

export type TextAutoFit = { type: 'none' } | { type: 'shape' } | { type: 'normal'; fontScale?: number; lineSpaceReduction?: number }

export interface Paragraph {
  runs: TextRun[]
  text: string
  style?: ParagraphStyle
}

export interface ParagraphStyle {
  align?: string
  level?: number
  indent?: number
  marginLeft?: number
  marginRight?: number
  defaultTabSize?: number
  rtl?: boolean
  bullet?: BulletStyle
  lineSpacing?: TextSpacing
  spaceBefore?: TextSpacing
  spaceAfter?: TextSpacing
  defaultRunStyle?: TextStyle
}

export interface BulletStyle {
  type: 'none' | 'character' | 'auto-number'
  character?: string
  autoNumberScheme?: string
  autoNumberStartAt?: number
  fontFace?: string
  color?: string
  fontSize?: number
}

export interface TextSpacing {
  points?: number
  percent?: number
}

export interface TextRun {
  text: string
  style?: TextStyle
}

export interface TextStyle {
  bold?: boolean
  italic?: boolean
  underline?: string
  fontSize?: number
  color?: string
  fontFace?: string
}

export interface SlideElementBase {
  id: string
  index: number
  type: SlideElement['type']
  name?: string
  transform?: Transform
  fill?: Fill
  line?: LineStyle
  geometry?: ShapeGeometry
  slidePart: string
  source: ElementSource
  visible: boolean
  opacity: number
  zIndex: number
  diagnostics?: Diagnostic[]
}

export interface TextElement extends SlideElementBase {
  type: 'text'
  text: string
  textBody: TextBody
}

export interface ImageCrop {
  left: number
  top: number
  right: number
  bottom: number
}

export interface ImageElement extends SlideElementBase {
  type: 'image'
  relationshipId?: string
  imagePart?: string
  crop?: ImageCrop
  image?: {
    part?: string
    isExternal: boolean
  }
  isExternal: boolean
}

export interface ShapeElement extends SlideElementBase {
  type: 'shape'
}

export interface ConnectorElement extends SlideElementBase {
  type: 'connector'
}

export interface UnknownElement extends SlideElementBase {
  type: 'unknown'
  nodeName: string
}
