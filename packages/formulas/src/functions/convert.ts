import { EMPTY, errorValue, numberValue, topLeftScalar } from '@ascend/schema'
import type { EvalArg, FunctionDef } from './registry.ts'
import { numArg } from './registry.ts'

interface UnitDef {
	readonly category: string
	readonly factor: number
}

const UNITS = new Map<string, UnitDef>([
	['g', { category: 'mass', factor: 1 }],
	['sg', { category: 'mass', factor: 14593.903 }],
	['lbm', { category: 'mass', factor: 453.59237 }],
	['u', { category: 'mass', factor: 1.6605390666e-24 }],
	['ozm', { category: 'mass', factor: 28.349523125 }],
	['grain', { category: 'mass', factor: 0.06479891 }],
	['cwt', { category: 'mass', factor: 45359.237 }],
	['shweight', { category: 'mass', factor: 45359.237 }],
	['uk_cwt', { category: 'mass', factor: 50802.34544 }],
	['stone', { category: 'mass', factor: 6350.29318 }],
	['ton', { category: 'mass', factor: 907184.74 }],
	['uk_ton', { category: 'mass', factor: 1016046.9088 }],
	['m', { category: 'distance', factor: 1 }],
	['mi', { category: 'distance', factor: 1609.344 }],
	['Nmi', { category: 'distance', factor: 1852 }],
	['in', { category: 'distance', factor: 0.0254 }],
	['ft', { category: 'distance', factor: 0.3048 }],
	['yd', { category: 'distance', factor: 0.9144 }],
	['ang', { category: 'distance', factor: 1e-10 }],
	['ell', { category: 'distance', factor: 1.143 }],
	['ly', { category: 'distance', factor: 9.4607304725808e15 }],
	['parsec', { category: 'distance', factor: 3.08567758149137e16 }],
	['pc', { category: 'distance', factor: 3.08567758149137e16 }],
	['Pica', { category: 'distance', factor: 0.0254 / 6 }],
	['picapt', { category: 'distance', factor: 0.0254 / 72 }],
	['pica', { category: 'distance', factor: 0.0254 / 72 }],
	['survey_mi', { category: 'distance', factor: 1609.3472186944 }],
	['yr', { category: 'time', factor: 365.25 * 86400 }],
	['day', { category: 'time', factor: 86400 }],
	['hr', { category: 'time', factor: 3600 }],
	['mn', { category: 'time', factor: 60 }],
	['min', { category: 'time', factor: 60 }],
	['sec', { category: 'time', factor: 1 }],
	['s', { category: 'time', factor: 1 }],
	['Pa', { category: 'pressure', factor: 1 }],
	['p', { category: 'pressure', factor: 1 }],
	['atm', { category: 'pressure', factor: 101325 }],
	['at', { category: 'pressure', factor: 101325 }],
	['mmHg', { category: 'pressure', factor: 133.322 }],
	['psi', { category: 'pressure', factor: 6894.757 }],
	['Torr', { category: 'pressure', factor: 133.3224 }],
	['N', { category: 'force', factor: 1 }],
	['dyn', { category: 'force', factor: 1e-5 }],
	['dy', { category: 'force', factor: 1e-5 }],
	['lbf', { category: 'force', factor: 4.44822162 }],
	['pond', { category: 'force', factor: 0.00980665 }],
	['J', { category: 'energy', factor: 1 }],
	['e', { category: 'energy', factor: 1e-7 }],
	['c', { category: 'energy', factor: 4.184 }],
	['cal', { category: 'energy', factor: 4.1868 }],
	['eV', { category: 'energy', factor: 1.602176634e-19 }],
	['ev', { category: 'energy', factor: 1.602176634e-19 }],
	['HPh', { category: 'energy', factor: 2684519.5 }],
	['hh', { category: 'energy', factor: 2684519.5 }],
	['Wh', { category: 'energy', factor: 3600 }],
	['wh', { category: 'energy', factor: 3600 }],
	['flb', { category: 'energy', factor: 1.3558179483 }],
	['BTU', { category: 'energy', factor: 1055.05585262 }],
	['btu', { category: 'energy', factor: 1055.05585262 }],
	['W', { category: 'power', factor: 1 }],
	['w', { category: 'power', factor: 1 }],
	['HP', { category: 'power', factor: 745.69987158227 }],
	['h', { category: 'power', factor: 745.69987158227 }],
	['PS', { category: 'power', factor: 735.49875 }],
	['T', { category: 'magnetism', factor: 1 }],
	['ga', { category: 'magnetism', factor: 1e-4 }],
	['C', { category: 'temperature', factor: 0 }],
	['F', { category: 'temperature', factor: 0 }],
	['K', { category: 'temperature', factor: 0 }],
	['kel', { category: 'temperature', factor: 0 }],
	['Rank', { category: 'temperature', factor: 0 }],
	['Reau', { category: 'temperature', factor: 0 }],
	['l', { category: 'volume', factor: 1 }],
	['L', { category: 'volume', factor: 1 }],
	['lt', { category: 'volume', factor: 1 }],
	['tsp', { category: 'volume', factor: 0.00492892159375 }],
	['tspm', { category: 'volume', factor: 0.005 }],
	['tbs', { category: 'volume', factor: 0.01478676478125 }],
	['oz', { category: 'volume', factor: 0.0295735295625 }],
	['cup', { category: 'volume', factor: 0.2365882365 }],
	['pt', { category: 'volume', factor: 0.473176473 }],
	['us_pt', { category: 'volume', factor: 0.473176473 }],
	['qt', { category: 'volume', factor: 0.946352946 }],
	['gal', { category: 'volume', factor: 3.785411784 }],
	['uk_pt', { category: 'volume', factor: 0.56826125 }],
	['uk_qt', { category: 'volume', factor: 1.1365225 }],
	['uk_gal', { category: 'volume', factor: 4.54609 }],
	['m3', { category: 'volume', factor: 1000 }],
	['in3', { category: 'volume', factor: 0.016387064 }],
	['ft3', { category: 'volume', factor: 28.316846592 }],
	['yd3', { category: 'volume', factor: 764.554857984 }],
	['barrel', { category: 'volume', factor: 158.987294928 }],
	['bushel', { category: 'volume', factor: 35.23907016688 }],
	['GRT', { category: 'volume', factor: 2831.6846592 }],
	['regton', { category: 'volume', factor: 2831.6846592 }],
	['MTON', { category: 'volume', factor: 1000 }],
	['m2', { category: 'area', factor: 1 }],
	['mi2', { category: 'area', factor: 2589988.110336 }],
	['in2', { category: 'area', factor: 0.00064516 }],
	['ft2', { category: 'area', factor: 0.09290304 }],
	['yd2', { category: 'area', factor: 0.83612736 }],
	['Nmi2', { category: 'area', factor: 3429904 }],
	['Morgen', { category: 'area', factor: 2500 }],
	['ar', { category: 'area', factor: 100 }],
	['ha', { category: 'area', factor: 10000 }],
	['uk_acre', { category: 'area', factor: 4046.8564224 }],
	['us_acre', { category: 'area', factor: 4046.8564224 }],
	['bit', { category: 'information', factor: 1 }],
	['byte', { category: 'information', factor: 8 }],
	['m/s', { category: 'speed', factor: 1 }],
	['m/h', { category: 'speed', factor: 1 / 3600 }],
	['mph', { category: 'speed', factor: 0.44704 }],
	['kn', { category: 'speed', factor: 1852 / 3600 }],
	['admkn', { category: 'speed', factor: 1853.184 / 3600 }],
])

const SI_PREFIXES: [string, number][] = [
	['Y', 1e24],
	['Z', 1e21],
	['E', 1e18],
	['P', 1e15],
	['T', 1e12],
	['G', 1e9],
	['M', 1e6],
	['k', 1e3],
	['h', 1e2],
	['da', 1e1],
	['d', 1e-1],
	['c', 1e-2],
	['m', 1e-3],
	['u', 1e-6],
	['n', 1e-9],
	['p', 1e-12],
	['f', 1e-15],
	['a', 1e-18],
	['z', 1e-21],
	['y', 1e-24],
]

const BINARY_PREFIXES: [string, number][] = [
	['Yi', 2 ** 80],
	['Zi', 2 ** 70],
	['Ei', 2 ** 60],
	['Pi', 2 ** 50],
	['Ti', 2 ** 40],
	['Gi', 2 ** 30],
	['Mi', 2 ** 20],
	['ki', 2 ** 10],
]

function resolveUnit(unit: string): { category: string; factor: number } | null {
	const direct = UNITS.get(unit)
	if (direct) return direct

	for (const [prefix, mult] of BINARY_PREFIXES) {
		if (unit.startsWith(prefix)) {
			const base = UNITS.get(unit.slice(prefix.length))
			if (base && base.category === 'information') {
				return { category: base.category, factor: base.factor * mult }
			}
		}
	}

	for (const [prefix, mult] of SI_PREFIXES) {
		if (unit.startsWith(prefix)) {
			const base = UNITS.get(unit.slice(prefix.length))
			if (base && base.category !== 'temperature') {
				return { category: base.category, factor: base.factor * mult }
			}
		}
	}

	return null
}

function toKelvin(value: number, unit: string): number {
	switch (unit) {
		case 'C':
			return value + 273.15
		case 'F':
			return ((value - 32) * 5) / 9 + 273.15
		case 'K':
		case 'kel':
			return value
		case 'Rank':
			return (value * 5) / 9
		case 'Reau':
			return (value * 5) / 4 + 273.15
		default:
			return value
	}
}

function fromKelvin(kelvin: number, unit: string): number {
	switch (unit) {
		case 'C':
			return kelvin - 273.15
		case 'F':
			return ((kelvin - 273.15) * 9) / 5 + 32
		case 'K':
		case 'kel':
			return kelvin
		case 'Rank':
			return (kelvin * 9) / 5
		case 'Reau':
			return ((kelvin - 273.15) * 4) / 5
		default:
			return kelvin
	}
}

export const convertFunction: FunctionDef = {
	name: 'CONVERT',
	minArgs: 3,
	maxArgs: 3,
	evaluate(args: EvalArg[]) {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		const fromVal = topLeftScalar(args[1]?.value ?? EMPTY)
		if (fromVal.kind !== 'string') return errorValue('#VALUE!')
		const toVal = topLeftScalar(args[2]?.value ?? EMPTY)
		if (toVal.kind !== 'string') return errorValue('#VALUE!')
		const fromUnit = fromVal.value
		const toUnit = toVal.value

		const from = resolveUnit(fromUnit)
		const to = resolveUnit(toUnit)
		if (!from || !to) return errorValue('#N/A')
		if (from.category !== to.category) return errorValue('#N/A')

		if (from.category === 'temperature') {
			const kelvin = toKelvin(n, fromUnit)
			return numberValue(fromKelvin(kelvin, toUnit))
		}

		return numberValue((n * from.factor) / to.factor)
	},
}
