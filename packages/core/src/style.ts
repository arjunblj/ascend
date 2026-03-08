export type Color =
	| { readonly kind: 'theme'; readonly theme: number; readonly tint?: number }
	| { readonly kind: 'rgb'; readonly rgb: string }
	| { readonly kind: 'indexed'; readonly index: number }
	| { readonly kind: 'auto' }

export type HorizontalAlign =
	| 'general'
	| 'left'
	| 'center'
	| 'right'
	| 'fill'
	| 'justify'
	| 'centerContinuous'
	| 'distributed'

export type VerticalAlign = 'top' | 'center' | 'bottom' | 'justify' | 'distributed'

export type FillPattern =
	| 'none'
	| 'solid'
	| 'darkGray'
	| 'mediumGray'
	| 'lightGray'
	| 'gray125'
	| 'gray0625'
	| 'darkHorizontal'
	| 'darkVertical'
	| 'darkDown'
	| 'darkUp'
	| 'darkGrid'
	| 'darkTrellis'
	| 'lightHorizontal'
	| 'lightVertical'
	| 'lightDown'
	| 'lightUp'
	| 'lightGrid'
	| 'lightTrellis'

export type BorderLineStyle =
	| 'none'
	| 'thin'
	| 'medium'
	| 'dashed'
	| 'dotted'
	| 'thick'
	| 'double'
	| 'hair'
	| 'mediumDashed'
	| 'dashDot'
	| 'mediumDashDot'
	| 'dashDotDot'
	| 'mediumDashDotDot'
	| 'slantDashDot'

export interface FontStyle {
	readonly name?: string
	readonly size?: number
	readonly bold?: boolean
	readonly italic?: boolean
	readonly underline?: boolean | 'single' | 'double'
	readonly strikethrough?: boolean
	readonly color?: Color
}

export interface FillStyle {
	readonly pattern?: FillPattern
	readonly fgColor?: Color
	readonly bgColor?: Color
}

export interface BorderEdge {
	readonly style?: BorderLineStyle
	readonly color?: Color
}

export interface BorderStyle {
	readonly top?: BorderEdge
	readonly bottom?: BorderEdge
	readonly left?: BorderEdge
	readonly right?: BorderEdge
	readonly diagonal?: BorderEdge
	readonly diagonalUp?: boolean
	readonly diagonalDown?: boolean
}

export interface AlignmentStyle {
	readonly horizontal?: HorizontalAlign
	readonly vertical?: VerticalAlign
	readonly wrapText?: boolean
	readonly shrinkToFit?: boolean
	readonly textRotation?: number
	readonly indent?: number
	readonly readingOrder?: number
}

export interface CellStyle {
	readonly font?: FontStyle
	readonly fill?: FillStyle
	readonly border?: BorderStyle
	readonly alignment?: AlignmentStyle
	readonly numberFormat?: string
	readonly protection?: {
		readonly locked?: boolean
		readonly hidden?: boolean
	}
}
