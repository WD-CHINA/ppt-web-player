import JSZip from 'jszip'

import type { Diagnostic } from '../diagnostics/Diagnostic'
import { createDiagnostic } from '../diagnostics/createDiagnostic'
import { parseXml } from '../xml/XmlParser'
import type { XmlNode } from '../xml/XmlNode'
import { parseContentTypes, type ContentTypes } from './ContentTypes'
import { normalizePartName, relationshipPartPath } from './PartName'
import type { Relationship, ResolvedRelationship } from './Relationship'
import { resolveRelationship } from './RelationshipResolver'
import * as xml from '../xml/XmlQuery'

export class PptxPackage {
  private readonly xmlCache = new Map<string, XmlNode | null>()
  private readonly relationshipsCache = new Map<string, Relationship[]>()
  private contentTypesCache: ContentTypes | null = null

  private constructor(
    private readonly zip: JSZip,
    private readonly diagnostics: Diagnostic[],
  ) {}

  static async load(input: Blob | ArrayBuffer): Promise<PptxPackage> {
    const diagnostics: Diagnostic[] = []
    const zip = await JSZip.loadAsync(input).catch((error: unknown) => {
      diagnostics.push(
        createDiagnostic({
          code: 'PPTX_ZIP_OPEN_FAILED',
          severity: 'error',
          message: '无法打开 PPTX ZIP 包。',
          detail: error,
        }),
      )
      return null
    })

    if (!zip) {
      return new PptxPackage(new JSZip(), diagnostics)
    }

    return new PptxPackage(zip, diagnostics)
  }

  getDiagnostics(): Diagnostic[] {
    return this.diagnostics
  }

  async getPart(path: string): Promise<ArrayBuffer | null> {
    const partPath = normalizePartName(path)
    const file = this.zip.file(partPath)

    if (!file) {
      this.diagnostics.push(
        createDiagnostic({
          code: 'PART_NOT_FOUND',
          severity: 'warning',
          part: partPath,
          message: `PPTX part 不存在：${partPath}`,
        }),
      )
      return null
    }

    return file.async('arraybuffer')
  }

  async getText(path: string): Promise<string | null> {
    const partPath = normalizePartName(path)
    const file = this.zip.file(partPath)

    if (!file) {
      this.diagnostics.push(
        createDiagnostic({
          code: 'PART_NOT_FOUND',
          severity: 'warning',
          part: partPath,
          message: `PPTX part 不存在：${partPath}`,
        }),
      )
      return null
    }

    return file.async('text')
  }

  async getXml(path: string): Promise<XmlNode | null> {
    const partPath = normalizePartName(path)

    if (this.xmlCache.has(partPath)) {
      return this.xmlCache.get(partPath) ?? null
    }

    const text = await this.getText(partPath)
    if (text === null) {
      this.xmlCache.set(partPath, null)
      return null
    }

    try {
      const node = parseXml(text)
      this.xmlCache.set(partPath, node)
      return node
    } catch (error) {
      this.diagnostics.push(
        createDiagnostic({
          code: 'XML_PARSE_FAILED',
          severity: 'error',
          part: partPath,
          message: `XML 解析失败：${partPath}`,
          detail: error,
        }),
      )
      this.xmlCache.set(partPath, null)
      return null
    }
  }

  async getContentTypes(): Promise<ContentTypes> {
    if (this.contentTypesCache) {
      return this.contentTypesCache
    }

    const root = await this.getXml('[Content_Types].xml')
    this.contentTypesCache = parseContentTypes(root)
    return this.contentTypesCache
  }

  async getRelationships(partPath: string): Promise<Relationship[]> {
    const normalizedPartPath = normalizePartName(partPath)

    if (this.relationshipsCache.has(normalizedPartPath)) {
      return this.relationshipsCache.get(normalizedPartPath) ?? []
    }

    const relsPath = relationshipPartPath(normalizedPartPath)
    const root = await this.getXml(relsPath)

    if (!root) {
      this.relationshipsCache.set(normalizedPartPath, [])
      return []
    }

    const relationships = xml.children(root, 'Relationship').flatMap((node) => {
      const id = xml.attr(node, 'Id')
      const type = xml.attr(node, 'Type')
      const target = xml.attr(node, 'Target')

      if (!id || !type || !target) {
        return []
      }

      return [
        {
          id,
          type,
          target,
          targetMode: xml.attr(node, 'TargetMode'),
        },
      ]
    })

    this.relationshipsCache.set(normalizedPartPath, relationships)
    return relationships
  }

  async resolveRelationship(partPath: string, rId: string): Promise<ResolvedRelationship | null> {
    const relationships = await this.getRelationships(partPath)
    const relationship = relationships.find((item) => item.id === rId)

    if (!relationship) {
      this.diagnostics.push(
        createDiagnostic({
          code: 'RELATIONSHIP_NOT_FOUND',
          severity: 'warning',
          part: partPath,
          message: `Relationship 不存在：${rId}`,
        }),
      )
      return null
    }

    return resolveRelationship(partPath, relationship)
  }
}
