import type { Diagnostic, DiagnosticSeverity } from './Diagnostic'
import type { DiagnosticCode } from './codes'

export interface DiagnosticInput {
  code: DiagnosticCode
  severity?: DiagnosticSeverity
  slideIndex?: number
  elementId?: string
  part?: string
  message: string
  detail?: unknown
}

export function createDiagnostic(input: DiagnosticInput): Diagnostic {
  return {
    code: input.code,
    severity: input.severity ?? 'warning',
    slideIndex: input.slideIndex,
    elementId: input.elementId,
    part: input.part,
    message: input.message,
    detail: input.detail,
  }
}
