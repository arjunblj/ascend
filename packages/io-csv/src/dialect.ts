import { type CsvDialect, DEFAULT_CSV_DIALECT } from '@ascend/schema'

export function resolveDialect(partial?: Partial<CsvDialect>): CsvDialect {
	if (!partial) return DEFAULT_CSV_DIALECT
	return { ...DEFAULT_CSV_DIALECT, ...partial }
}

export const TSV_DIALECT: Partial<CsvDialect> = {
	delimiter: '\t',
	hasHeader: true,
}
