import { describe, expect, test } from 'bun:test'
import { fileURLToPath } from 'node:url'

const runnerPath = fileURLToPath(new URL('./formula-sota.ts', import.meta.url))

const profiles = [
	{
		name: 'hf-prefix-range-sum',
		args: ['--rows', '120', '--formulas', '120'],
		aggregates: ['SUM', 'COUNT', 'AVERAGE', 'MIN', 'MAX'],
	},
	{
		name: 'hf-prefix-range-dirty-head',
		args: ['--rows', '120', '--formulas', '120'],
		aggregates: ['SUM', 'COUNT', 'AVERAGE', 'MIN', 'MAX'],
	},
	{
		name: 'hf-prefix-range-dirty-tail',
		args: ['--rows', '120', '--formulas', '120'],
		aggregates: ['SUM', 'COUNT', 'AVERAGE', 'MIN', 'MAX'],
	},
	{
		name: 'hf-indexed-index-match',
		args: ['--rows', '180', '--formulas', '40'],
	},
	{
		name: 'hf-indexed-index-match-dirty-key',
		args: ['--rows', '180', '--formulas', '40'],
	},
	{
		name: 'hf-indexed-index-match-dirty-value',
		args: ['--rows', '180', '--formulas', '40'],
	},
] as const

interface FormulaSotaPayload {
	readonly suite: string
	readonly profile: {
		readonly name: string
		readonly sourceBenchmark: string
		readonly sourceUrl: string
		readonly aggregate?: string
	}
	readonly cases: readonly {
		readonly engine: string
		readonly correctness: Record<string, string | number | boolean>
	}[]
	readonly comparison: {
		readonly totalSpeedupVsHyperFormula: number
	}
}

describe('formula SOTA public profile smoke', () => {
	for (const profile of profiles) {
		const aggregateCases = 'aggregates' in profile ? profile.aggregates : [undefined]
		for (const aggregate of aggregateCases) {
			test(`${profile.name}${aggregate ? ` ${aggregate}` : ''} stays correct against the public comparator shape`, () => {
				const aggregateArgs = aggregate ? ['--aggregate', aggregate] : []
				const proc = Bun.spawnSync({
					cmd: [
						Bun.argv[0],
						runnerPath,
						'--profile',
						profile.name,
						...profile.args,
						...aggregateArgs,
						'--repeat',
						'1',
						'--warmup',
						'0',
						'--json',
					],
					stdout: 'pipe',
					stderr: 'pipe',
				})
				const stderr = new TextDecoder().decode(proc.stderr)
				expect(proc.exitCode, stderr).toBe(0)

				const payload = JSON.parse(new TextDecoder().decode(proc.stdout)) as FormulaSotaPayload
				expect(payload.suite).toBe('ascend-formula-sota')
				expect(payload.profile.name).toBe(profile.name)
				if (aggregate) expect(payload.profile.aggregate).toBe(aggregate)
				expect(payload.profile.sourceBenchmark.length).toBeGreaterThan(0)
				expect(payload.profile.sourceUrl).toContain('hyperformula')
				expect(payload.cases.map((entry) => entry.engine).sort()).toEqual([
					'ascend',
					'hyperformula',
				])
				expect(payload.comparison.totalSpeedupVsHyperFormula).toBeGreaterThan(1)
				for (const entry of payload.cases) {
					const matchFlags = Object.entries(entry.correctness).filter(([key]) =>
						key.endsWith('Matches'),
					)
					expect(matchFlags.length).toBeGreaterThan(0)
					for (const [key, value] of matchFlags) {
						expect(value, `${entry.engine}.${key}`).toBe(true)
					}
					if (entry.engine === 'ascend') {
						expect(entry.correctness.errors ?? 0).toBe(0)
					}
				}
			})
		}
	}
})
