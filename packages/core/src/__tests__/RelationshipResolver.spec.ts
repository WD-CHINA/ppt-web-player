import { describe, expect, it } from 'vitest'

import { resolveTargetPath } from '../package/RelationshipResolver'

describe('RelationshipResolver', () => {
  it('resolves relative relationship targets from the source part directory', () => {
    expect(resolveTargetPath('ppt/slides/slide1.xml', '../media/image1.png')).toBe(
      'ppt/media/image1.png',
    )
  })

  it('normalizes absolute targets inside the package', () => {
    expect(resolveTargetPath('ppt/presentation.xml', '/ppt/slides/slide1.xml')).toBe(
      'ppt/slides/slide1.xml',
    )
  })
})
