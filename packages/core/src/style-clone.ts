import type { CellStyle } from './style.ts'

export function cloneCellStyle(style: CellStyle): CellStyle {
	return {
		...(style.font
			? { font: { ...style.font, ...(style.font.color ? { color: { ...style.font.color } } : {}) } }
			: {}),
		...(style.fill
			? {
					fill: {
						...style.fill,
						...(style.fill.fgColor ? { fgColor: { ...style.fill.fgColor } } : {}),
						...(style.fill.bgColor ? { bgColor: { ...style.fill.bgColor } } : {}),
					},
				}
			: {}),
		...(style.border
			? {
					border: {
						...style.border,
						...(style.border.top
							? {
									top: {
										...style.border.top,
										...(style.border.top.color ? { color: { ...style.border.top.color } } : {}),
									},
								}
							: {}),
						...(style.border.bottom
							? {
									bottom: {
										...style.border.bottom,
										...(style.border.bottom.color
											? { color: { ...style.border.bottom.color } }
											: {}),
									},
								}
							: {}),
						...(style.border.left
							? {
									left: {
										...style.border.left,
										...(style.border.left.color ? { color: { ...style.border.left.color } } : {}),
									},
								}
							: {}),
						...(style.border.right
							? {
									right: {
										...style.border.right,
										...(style.border.right.color ? { color: { ...style.border.right.color } } : {}),
									},
								}
							: {}),
						...(style.border.diagonal
							? {
									diagonal: {
										...style.border.diagonal,
										...(style.border.diagonal.color
											? { color: { ...style.border.diagonal.color } }
											: {}),
									},
								}
							: {}),
					},
				}
			: {}),
		...(style.alignment ? { alignment: { ...style.alignment } } : {}),
		...(style.numberFormat ? { numberFormat: style.numberFormat } : {}),
		...(style.protection ? { protection: { ...style.protection } } : {}),
	}
}
