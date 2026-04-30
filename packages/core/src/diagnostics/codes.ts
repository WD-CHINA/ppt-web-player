export const DIAGNOSTIC_CODES = {
  pptxZipOpenFailed: 'PPTX_ZIP_OPEN_FAILED',
  partNotFound: 'PART_NOT_FOUND',
  xmlParseFailed: 'XML_PARSE_FAILED',
  relationshipNotFound: 'RELATIONSHIP_NOT_FOUND',
  imageRelationshipNotFound: 'IMAGE_RELATIONSHIP_NOT_FOUND',
  imageRelationshipResolveFailed: 'IMAGE_RELATIONSHIP_RESOLVE_FAILED',
  imageCropInvalid: 'IMAGE_CROP_INVALID',
  imageAlphaInvalid: 'IMAGE_ALPHA_INVALID',
  presentationNotFound: 'PRESENTATION_NOT_FOUND',
  slideSizeNotFound: 'SLIDE_SIZE_NOT_FOUND',
  slideListNotFound: 'SLIDE_LIST_NOT_FOUND',
  slidePartNotFound: 'SLIDE_PART_NOT_FOUND',
  slideShapeTreeNotFound: 'SLIDE_SHAPE_TREE_NOT_FOUND',
  unexpectedRelationshipType: 'UNEXPECTED_RELATIONSHIP_TYPE',
  unsupportedSlideElement: 'UNSUPPORTED_SLIDE_ELEMENT',
  styleInheritanceIncomplete: 'STYLE_INHERITANCE_INCOMPLETE',
} as const

export type DiagnosticCode = (typeof DIAGNOSTIC_CODES)[keyof typeof DIAGNOSTIC_CODES]
