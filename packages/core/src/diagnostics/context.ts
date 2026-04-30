import { createDiagnostic, type DiagnosticInput } from './createDiagnostic'
import type { Diagnostic } from './Diagnostic'

export interface DiagnosticContext {
  part?: string
  slideIndex?: number
  elementId?: string
}

export function withDiagnosticContext(
  input: DiagnosticInput,
  context: DiagnosticContext | undefined,
): Diagnostic {
  return createDiagnostic({
    ...input,
    part: input.part ?? context?.part,
    slideIndex: input.slideIndex ?? context?.slideIndex,
    elementId: input.elementId ?? context?.elementId,
  })
}

export function pushDiagnostic(
  diagnostics: Diagnostic[],
  input: DiagnosticInput,
  context?: DiagnosticContext,
): Diagnostic {
  const diagnostic = withDiagnosticContext(input, context)
  diagnostics.push(diagnostic)
  return diagnostic
}

export function elementDiagnosticContext(
  part: string,
  slideIndex: number | undefined,
  elementId: string,
): DiagnosticContext {
  return {
    part,
    slideIndex,
    elementId,
  }
}

export function slideDiagnosticContext(part: string, slideIndex?: number): DiagnosticContext {
  return {
    part,
    slideIndex,
  }
}
