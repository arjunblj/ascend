export enum TokenType {
	Number = 'Number',
	String = 'String',
	Boolean = 'Boolean',
	Error = 'Error',
	CellRef = 'CellRef',
	Name = 'Name',
	Function = 'Function',
	Operator = 'Operator',
	OpenParen = 'OpenParen',
	CloseParen = 'CloseParen',
	OpenBrace = 'OpenBrace',
	CloseBrace = 'CloseBrace',
	Comma = 'Comma',
	Semicolon = 'Semicolon',
	Colon = 'Colon',
	Bang = 'Bang',
	Whitespace = 'Whitespace',
	EOF = 'EOF',
}

export interface Token {
	readonly type: TokenType
	readonly value: string
	readonly position: number
}
