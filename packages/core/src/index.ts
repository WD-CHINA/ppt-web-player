export type { Diagnostic, DiagnosticSeverity } from './diagnostics/Diagnostic'
export type {
  ConnectorElement,
  Fill,
  ImageElement,
  LineEndStyle,
  LineStyle,
  NoFill,
  Presentation,
  PresentationMetadata,
  PresetShapeGeometry,
  ShapeElement,
  ShapeGeometry,
  Slide,
  SlideElement,
  SolidFill,
  TextElement,
  Transform,
  UnknownElement,
} from './model/Presentation'
export type { Relationship, ResolvedRelationship } from './package/Relationship'
export { parsePptx, type ParsePptxResult } from './parser/parsePptx'
export { PptxPackage } from './package/PptxPackage'
