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

export interface SlideElementBase {
  id: string
  index: number
  name?: string
  transform?: Transform
  fill?: Fill
  line?: LineStyle
  geometry?: ShapeGeometry
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
