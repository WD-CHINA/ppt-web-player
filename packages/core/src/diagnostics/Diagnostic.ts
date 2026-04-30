import type { DiagnosticCode } from './codes'

export type DiagnosticSeverity = 'info' | 'warning' | 'error'

export interface Diagnostic {
  code: DiagnosticCode
  severity: DiagnosticSeverity
  slideIndex?: number
  elementId?: string
  part?: string
  message: string
  detail?: unknown
}
