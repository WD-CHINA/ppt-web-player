declare module 'txml' {
  export interface tNode {
    tagName: string
    attributes: Record<string, string | number | boolean>
    children: Array<tNode | string>
  }

  export interface TParseOptions {
    pos?: number
    noChildNodes?: string[]
    setPos?: boolean
    keepComments?: boolean
    keepWhitespace?: boolean
    simplify?: boolean
    filter?: (a: tNode, b: tNode) => boolean
  }

  export function parse(input: string, options?: TParseOptions): Array<tNode | string>
}
