export type DiagnosticSeverity = 'info' | 'warning' | 'error'

export interface Diagnostic {
  code: string
  severity: DiagnosticSeverity
  slideIndex?: number
  elementId?: string
  part?: string
  message: string
  detail?: unknown
}
