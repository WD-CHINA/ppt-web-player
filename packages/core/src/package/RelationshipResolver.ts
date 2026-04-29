import { dirname, normalizePartName } from './PartName'
import type { Relationship, ResolvedRelationship } from './Relationship'

export function resolveRelationship(basePart: string, relationship: Relationship): ResolvedRelationship {
  const isExternal = relationship.targetMode === 'External' || /^[a-z][a-z0-9+.-]*:/i.test(relationship.target)

  return {
    ...relationship,
    path: isExternal ? relationship.target : resolveTargetPath(basePart, relationship.target),
    isExternal,
  }
}

export function resolveTargetPath(basePart: string, target: string): string {
  if (target.startsWith('/')) {
    return normalizePartName(target)
  }

  return normalizePartName(`${dirname(basePart)}/${target}`)
}
