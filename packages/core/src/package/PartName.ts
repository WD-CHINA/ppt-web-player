export function normalizePartName(path: string): string {
  const cleanPath = path.replace(/\\/g, '/').replace(/^\//, '')
  const segments: string[] = []

  for (const segment of cleanPath.split('/')) {
    if (!segment || segment === '.') {
      continue
    }

    if (segment === '..') {
      segments.pop()
      continue
    }

    segments.push(segment)
  }

  return segments.join('/')
}

export function dirname(partPath: string): string {
  const normalized = normalizePartName(partPath)
  const index = normalized.lastIndexOf('/')
  return index === -1 ? '' : normalized.slice(0, index)
}

export function relationshipPartPath(partPath: string): string {
  if (partPath === '') {
    return '_rels/.rels'
  }

  const base = dirname(partPath)
  const segments = normalizePartName(partPath).split('/')
  const fileName = segments[segments.length - 1]
  return normalizePartName(`${base}/_rels/${fileName}.rels`)
}
