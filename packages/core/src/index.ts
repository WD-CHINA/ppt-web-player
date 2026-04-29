export type { Diagnostic, DiagnosticSeverity } from './diagnostics/Diagnostic'
export type {
  ConnectorElement,
  Fill,
  ImageElement,
  LineStyle,
  Presentation,
  PresentationMetadata,
  ShapeElement,
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
