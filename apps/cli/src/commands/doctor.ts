import { jsonOut } from '../output/json.ts'

export const usage = `Usage: ascend doctor [flags]

  Check environment and verify dependencies.

Flags:
  --json          Output as JSON
`

interface Check {
	name: string
	status: 'ok' | 'warn' | 'fail'
	detail: string
}

export async function doctorCommand(_args: string[], flags: Map<string, string>): Promise<number> {
	const checks: Check[] = []

	const bunVersion = typeof Bun !== 'undefined' ? Bun.version : null
	if (bunVersion) {
		checks.push({ name: 'bun', status: 'ok', detail: `v${bunVersion}` })
	} else {
		checks.push({ name: 'bun', status: 'fail', detail: 'Bun runtime not detected' })
	}

	const nodeVersion = process.version
	checks.push({ name: 'node-compat', status: 'ok', detail: nodeVersion })

	try {
		await import('@ascend/sdk')
		checks.push({ name: '@ascend/sdk', status: 'ok', detail: 'loaded' })
	} catch {
		checks.push({ name: '@ascend/sdk', status: 'fail', detail: 'failed to load' })
	}

	if (flags.has('json')) {
		console.log(jsonOut({ checks }))
	} else {
		console.log('ascend doctor\n')
		for (const c of checks) {
			const icon = c.status === 'ok' ? '+' : c.status === 'warn' ? '~' : 'x'
			console.log(`  [${icon}] ${c.name}: ${c.detail}`)
		}
	}

	const failed = checks.some((c) => c.status === 'fail')
	return failed ? 1 : 0
}
