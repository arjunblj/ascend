/**
 * Audit: list commonly-used Excel functions NOT in Ascend's registry.
 * Run: bun run fixtures/formulas/missing-formula-audit.ts
 */
import { readdirSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { functionRegistry } from '../../packages/formulas/src/index.ts'

export const COMMON_EXCEL_FUNCTIONS = [
	'ABS',
	'ACOS',
	'ACOSH',
	'ACOT',
	'ACOTH',
	'ADDRESS',
	'AGGREGATE',
	'AND',
	'ARABIC',
	'AREAS',
	'ASIN',
	'ASINH',
	'ATAN',
	'ATAN2',
	'ATANH',
	'AVEDEV',
	'AVERAGE',
	'AVERAGEA',
	'AVERAGEIF',
	'AVERAGEIFS',
	'BASE',
	'BESSELI',
	'BESSELJ',
	'BESSELK',
	'BESSELY',
	'BETA.DIST',
	'BETA.INV',
	'BIN2DEC',
	'BIN2HEX',
	'BIN2OCT',
	'BINOM.DIST',
	'BINOM.DIST.RANGE',
	'BINOM.INV',
	'BITAND',
	'BITLSHIFT',
	'BITOR',
	'BITRSHIFT',
	'BITXOR',
	'BYCOL',
	'BYROW',
	'CEILING',
	'CEILING.MATH',
	'CEILING.PRECISE',
	'CELL',
	'CHAR',
	'CHISQ.DIST',
	'CHISQ.DIST.RT',
	'CHISQ.INV',
	'CHISQ.INV.RT',
	'CHISQ.TEST',
	'CHOOSE',
	'CHOOSECOLS',
	'CHOOSEROWS',
	'CLEAN',
	'CODE',
	'COLUMN',
	'COLUMNS',
	'COMBIN',
	'COMBINA',
	'COMPLEX',
	'CONCAT',
	'CONCATENATE',
	'CONFIDENCE.NORM',
	'CONFIDENCE.T',
	'CONVERT',
	'CORREL',
	'COS',
	'COSH',
	'COT',
	'COTH',
	'COUNT',
	'COUNTA',
	'COUNTBLANK',
	'COUNTIF',
	'COUNTIFS',
	'COVARIANCE.P',
	'COVARIANCE.S',
	'CSC',
	'CSCH',
	'DATE',
	'DATEDIF',
	'DATEVALUE',
	'DAVERAGE',
	'DAY',
	'DAYS',
	'DAYS360',
	'DB',
	'DCOUNT',
	'DCOUNTA',
	'DDB',
	'DEC2BIN',
	'DEC2HEX',
	'DEC2OCT',
	'DECIMAL',
	'DEGREES',
	'DELTA',
	'DEVSQ',
	'DGET',
	'DISC',
	'DMAX',
	'DMIN',
	'DOLLAR',
	'DROP',
	'DSTDEV',
	'DSTDEVP',
	'DSUM',
	'DVAR',
	'DVARP',
	'EDATE',
	'EFFECT',
	'EOMONTH',
	'ERF',
	'ERF.PRECISE',
	'ERFC',
	'ERFC.PRECISE',
	'ERROR.TYPE',
	'EVEN',
	'EXACT',
	'EXP',
	'EXPAND',
	'EXPON.DIST',
	'F.DIST',
	'F.DIST.RT',
	'F.INV',
	'F.INV.RT',
	'F.TEST',
	'FACT',
	'FACTDOUBLE',
	'FALSE',
	'FILTER',
	'FIND',
	'FISHER',
	'FISHERINV',
	'FIXED',
	'FLOOR',
	'FLOOR.MATH',
	'FLOOR.PRECISE',
	'FORECAST',
	'FORECAST.ETS',
	'FORECAST.ETS.CONFINT',
	'FORECAST.ETS.SEASONALITY',
	'FORECAST.ETS.STAT',
	'FORECAST.LINEAR',
	'FORMULATEXT',
	'FREQUENCY',
	'FV',
	'FVSCHEDULE',
	'GAMMA',
	'GAMMA.DIST',
	'GAMMA.INV',
	'GAMMALN',
	'GAMMALN.PRECISE',
	'GCD',
	'GEOMEAN',
	'GESTEP',
	'GROWTH',
	'HARMEAN',
	'HEX2BIN',
	'HEX2DEC',
	'HEX2OCT',
	'HLOOKUP',
	'HOUR',
	'HSTACK',
	'HYPERLINK',
	'IF',
	'IFERROR',
	'IFNA',
	'IFS',
	'IMABS',
	'IMAGINARY',
	'IMARGUMENT',
	'IMCONJUGATE',
	'IMCOS',
	'IMDIV',
	'IMEXP',
	'IMLN',
	'IMLOG10',
	'IMLOG2',
	'IMPOWER',
	'IMPRODUCT',
	'IMREAL',
	'IMSIN',
	'IMSQRT',
	'IMSUB',
	'IMSUM',
	'INDEX',
	'INDIRECT',
	'INT',
	'INTERCEPT',
	'INTRATE',
	'IPMT',
	'IRR',
	'ISBLANK',
	'ISERR',
	'ISERROR',
	'ISEVEN',
	'ISFORMULA',
	'ISLOGICAL',
	'ISNA',
	'ISNONTEXT',
	'ISNUMBER',
	'ISODD',
	'ISOWEEKNUM',
	'ISPMT',
	'ISREF',
	'ISTEXT',
	'KURT',
	'LAMBDA',
	'LARGE',
	'LCM',
	'LEFT',
	'LEN',
	'LET',
	'LINEST',
	'LN',
	'LOG',
	'LOG10',
	'LOGEST',
	'LOGNORM.DIST',
	'LOGNORM.INV',
	'LOOKUP',
	'LOWER',
	'MAKEARRAY',
	'MAP',
	'MATCH',
	'MAX',
	'MAXA',
	'MAXIFS',
	'MDETERM',
	'MEDIAN',
	'MID',
	'MIN',
	'MINA',
	'MINIFS',
	'MINUTE',
	'MINVERSE',
	'MIRR',
	'MMULT',
	'MOD',
	'MODE',
	'MODE.MULT',
	'MODE.SNGL',
	'MONTH',
	'MROUND',
	'MULTINOMIAL',
	'MUNIT',
	'N',
	'NA',
	'NETWORKDAYS',
	'NETWORKDAYS.INTL',
	'NOMINAL',
	'NORM.DIST',
	'NORM.INV',
	'NORM.S.DIST',
	'NORM.S.INV',
	'NOT',
	'NOW',
	'NPER',
	'NPV',
	'NUMBERVALUE',
	'OCT2BIN',
	'OCT2DEC',
	'OCT2HEX',
	'ODD',
	'OFFSET',
	'OR',
	'PDURATION',
	'PEARSON',
	'PERCENTILE',
	'PERCENTILE.EXC',
	'PERCENTILE.INC',
	'PERCENTRANK',
	'PERCENTRANK.EXC',
	'PERCENTRANK.INC',
	'PERMUT',
	'PERMUTATIONA',
	'PHI',
	'PI',
	'PMT',
	'POISSON.DIST',
	'POWER',
	'PPMT',
	'PRICE',
	'PRICEDISC',
	'PRICEMAT',
	'PRODUCT',
	'PROPER',
	'PV',
	'QUARTILE',
	'QUARTILE.EXC',
	'QUARTILE.INC',
	'QUOTIENT',
	'RADIANS',
	'RAND',
	'RANDARRAY',
	'RANDBETWEEN',
	'RANK',
	'RANK.AVG',
	'RANK.EQ',
	'RATE',
	'RECEIVED',
	'REDUCE',
	'REPLACE',
	'REPT',
	'RIGHT',
	'ROMAN',
	'ROUND',
	'ROUNDDOWN',
	'ROUNDUP',
	'ROW',
	'ROWS',
	'RRI',
	'RSQ',
	'SCAN',
	'SEARCH',
	'SEC',
	'SECH',
	'SECOND',
	'SEQUENCE',
	'SERIESSUM',
	'SHEET',
	'SHEETS',
	'SIGN',
	'SIN',
	'SINH',
	'SKEW',
	'SKEW.P',
	'SLN',
	'SLOPE',
	'SMALL',
	'SORT',
	'SORTBY',
	'SQRT',
	'SQRTPI',
	'STANDARDIZE',
	'STDEV',
	'STDEV.P',
	'STDEV.S',
	'STDEVA',
	'STDEVPA',
	'STEYX',
	'SUBSTITUTE',
	'SUBTOTAL',
	'SUM',
	'SUMIF',
	'SUMIFS',
	'SUMPRODUCT',
	'SUMSQ',
	'SUMX2MY2',
	'SUMX2PY2',
	'SUMXMY2',
	'SWITCH',
	'SYD',
	'T',
	'T.DIST',
	'T.DIST.2T',
	'T.DIST.RT',
	'T.INV',
	'T.INV.2T',
	'T.TEST',
	'TAKE',
	'TAN',
	'TANH',
	'TBILLEQ',
	'TBILLPRICE',
	'TBILLYIELD',
	'TEXT',
	'TEXTAFTER',
	'TEXTBEFORE',
	'TEXTJOIN',
	'TEXTSPLIT',
	'TIME',
	'TIMEVALUE',
	'TOCOL',
	'TODAY',
	'TOROW',
	'TRANSPOSE',
	'TREND',
	'TRIM',
	'TRUE',
	'TRUNC',
	'TYPE',
	'UNICHAR',
	'UNICODE',
	'UNIQUE',
	'UPPER',
	'VALUE',
	'VALUETOTEXT',
	'VAR',
	'VAR.P',
	'VAR.S',
	'VARA',
	'VARPA',
	'VDB',
	'VLOOKUP',
	'VSTACK',
	'WEEKDAY',
	'WEEKNUM',
	'WEIBULL.DIST',
	'WORKDAY',
	'WORKDAY.INTL',
	'WRAPCOLS',
	'WRAPROWS',
	'XIRR',
	'XLOOKUP',
	'XMATCH',
	'XNPV',
	'XOR',
	'YEAR',
	'YEARFRAC',
	'YIELD',
	'YIELDDISC',
	'YIELDMAT',
]

export interface FormulaCorpusCoverage {
	readonly fixtureFiles: number
	readonly totalCases: number
	readonly coveredFunctions: string[]
	readonly trackedCovered: string[]
	readonly presentButNotCorpusCovered: string[]
	readonly trackedCoverage: number
}

export interface MissingFormulaAudit {
	readonly total: number
	readonly present: string[]
	readonly missing: string[]
	readonly coverage: number
	readonly semanticCorpus: FormulaCorpusCoverage
}

interface FormulaConformanceCase {
	readonly formula?: string
	readonly setupFormulas?: Readonly<Record<string, string>>
}

interface FormulaConformanceFixture {
	readonly cases?: readonly FormulaConformanceCase[]
}

const FORMULA_CALL_RE = /(?<![A-Z0-9_.])([A-Z_][A-Z0-9_.]*)\s*\(/gi

function normalizeFunctionName(name: string): string {
	return name
		.toUpperCase()
		.replace(/^_XLFN\./, '')
		.replace(/^_XLWS\./, '')
}

function stripQuotedStrings(formula: string): string {
	return formula.replace(/"(?:""|[^"])*"/g, '""')
}

function collectFormulaCalls(formula: string, out: Set<string>): void {
	const searchable = stripQuotedStrings(formula)
	for (const match of searchable.matchAll(FORMULA_CALL_RE)) {
		const name = match[1]
		if (name) out.add(normalizeFunctionName(name))
	}
}

export function collectFormulaCorpusCoverage(
	fixturesDir = import.meta.dir,
	trackedFunctions: readonly string[] = COMMON_EXCEL_FUNCTIONS,
): FormulaCorpusCoverage {
	const jsonFiles = readdirSync(fixturesDir)
		.filter((file) => file.endsWith('.json') && file !== 'package.json')
		.sort()
	const covered = new Set<string>()
	let totalCases = 0

	for (const file of jsonFiles) {
		const fixture = JSON.parse(
			readFileSync(join(fixturesDir, file), 'utf-8'),
		) as FormulaConformanceFixture
		for (const testCase of fixture.cases ?? []) {
			totalCases++
			if (testCase.formula) collectFormulaCalls(testCase.formula, covered)
			for (const formula of Object.values(testCase.setupFormulas ?? {})) {
				collectFormulaCalls(formula, covered)
			}
		}
	}

	const tracked = trackedFunctions.map(normalizeFunctionName).sort()
	const trackedCovered = tracked.filter((name) => covered.has(name))
	const presentButNotCorpusCovered = tracked.filter((name) => !covered.has(name))

	return {
		fixtureFiles: jsonFiles.length,
		totalCases,
		coveredFunctions: [...covered].sort(),
		trackedCovered,
		presentButNotCorpusCovered,
		trackedCoverage: trackedCovered.length / tracked.length,
	}
}

export function runMissingFormulaAudit(): MissingFormulaAudit {
	const registered = new Set<string>()
	for (const fn of functionRegistry.values()) {
		registered.add(fn.name.toUpperCase())
	}

	const missing: string[] = []
	const present: string[] = []
	for (const name of COMMON_EXCEL_FUNCTIONS) {
		if (registered.has(name.toUpperCase())) present.push(name)
		else missing.push(name)
	}

	return {
		total: COMMON_EXCEL_FUNCTIONS.length,
		present,
		missing,
		coverage: present.length / COMMON_EXCEL_FUNCTIONS.length,
		semanticCorpus: collectFormulaCorpusCoverage(),
	}
}

function main(): void {
	const { total, present, missing, coverage, semanticCorpus } = runMissingFormulaAudit()
	console.log('Missing Formula Functions Audit')
	console.log('='.repeat(60))
	console.log(`Total common functions checked: ${String(total)}`)
	console.log(`Present in Ascend: ${String(present.length)}`)
	console.log(`Missing from Ascend: ${String(missing.length)}`)
	console.log(`JSON conformance fixture files: ${String(semanticCorpus.fixtureFiles)}`)
	console.log(`JSON conformance cases: ${String(semanticCorpus.totalCases)}`)
	console.log(
		`Common functions covered by JSON conformance: ${String(semanticCorpus.trackedCovered.length)}`,
	)
	console.log(
		`Common functions present but not JSON-corpus covered: ${String(
			semanticCorpus.presentButNotCorpusCovered.length,
		)}`,
	)
	console.log()
	if (missing.length > 0) {
		console.log('Missing functions:')
		for (const name of missing) console.log(`  ${name}`)
	}
	if (semanticCorpus.presentButNotCorpusCovered.length > 0) {
		console.log('Present common functions without JSON conformance cases:')
		for (const name of semanticCorpus.presentButNotCorpusCovered) console.log(`  ${name}`)
		console.log()
	}
	console.log()
	console.log(`Coverage: ${(coverage * 100).toFixed(1)}%`)
	console.log(`JSON semantic coverage: ${(semanticCorpus.trackedCoverage * 100).toFixed(1)}%`)
	if (missing.length > 0) process.exit(1)
}

if (import.meta.main) {
	main()
}
