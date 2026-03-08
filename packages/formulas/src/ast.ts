import type { ExcelError } from '@ascend/schema'

export type BinaryOp =
	| '+'
	| '-'
	| '*'
	| '/'
	| '^'
	| '&'
	| '='
	| '<>'
	| '<'
	| '>'
	| '<='
	| '>='
	| ' '
	| ','

export type UnaryOp = '+' | '-' | '%'

export interface FormulaCellRef {
	readonly row: number
	readonly col: number
	readonly rowAbsolute: boolean
	readonly colAbsolute: boolean
}

export interface NumberNode {
	readonly type: 'number'
	readonly value: number
}

export interface StringNode {
	readonly type: 'string'
	readonly value: string
}

export interface BooleanNode {
	readonly type: 'boolean'
	readonly value: boolean
}

export interface ErrorNode {
	readonly type: 'error'
	readonly value: ExcelError
}

export interface CellRefNode {
	readonly type: 'cellRef'
	readonly ref: FormulaCellRef
	readonly sheet?: string
}

export interface RangeRefNode {
	readonly type: 'rangeRef'
	readonly start: FormulaCellRef
	readonly end: FormulaCellRef
	readonly sheet?: string
}

export interface NameNode {
	readonly type: 'name'
	readonly name: string
}

export interface FunctionNode {
	readonly type: 'function'
	readonly name: string
	readonly args: readonly FormulaNode[]
}

export interface BinaryNode {
	readonly type: 'binary'
	readonly op: BinaryOp
	readonly left: FormulaNode
	readonly right: FormulaNode
}

export interface UnaryNode {
	readonly type: 'unary'
	readonly op: UnaryOp
	readonly operand: FormulaNode
}

export interface ArrayNode {
	readonly type: 'array'
	readonly rows: readonly (readonly FormulaNode[])[]
}

export interface StructuredRefNode {
	readonly type: 'structuredRef'
	readonly table: string
	readonly specifiers: readonly string[]
	readonly column?: string
}

export interface MissingNode {
	readonly type: 'missing'
}

export type FormulaNode =
	| NumberNode
	| StringNode
	| BooleanNode
	| ErrorNode
	| CellRefNode
	| RangeRefNode
	| NameNode
	| FunctionNode
	| BinaryNode
	| UnaryNode
	| ArrayNode
	| StructuredRefNode
	| MissingNode
