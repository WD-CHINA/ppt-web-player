export interface Relationship {
  id: string
  type: string
  target: string
  targetMode?: string
}

export interface ResolvedRelationship extends Relationship {
  path: string
  isExternal: boolean
}
