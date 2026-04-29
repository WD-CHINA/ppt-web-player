import type { Diagnostic } from '../diagnostics/Diagnostic'

export interface Presentation {
  id: string
  width: number
  height: number
  slides: Slide[]
  diagnostics: Diagnostic[]
  metadata: PresentationMetadata
}

export interface PresentationMetadata {
  presentationPart: string
}

export interface Slide {
  id: string
  index: number
  part: string
  relationshipId: string
  name?: string
  background?: Fill
  elements: SlideElement[]
  diagnostics: Diagnostic[]
}

export type SlideElement = TextElement | ImageElement | ShapeElement | ConnectorElement | UnknownElement

export type Fill = SolidFill

export interface SolidFill {
  type: 'solid'
  color: string
}

export interface LineStyle {
  color?: string
  width?: number
}

export interface Transform {
  x: number
  y: number
  width: number
  height: number
}

export interface SlideElementBase {
  id: string
  index: number
  name?: string
  transform?: Transform
  fill?: Fill
  line?: LineStyle
}

export interface TextElement extends SlideElementBase {
  type: 'text'
  text: string
}

export interface ImageElement extends SlideElementBase {
  type: 'image'
  relationshipId?: string
  part?: string
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
