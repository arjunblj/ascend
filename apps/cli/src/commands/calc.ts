import { AscendWorkbook } from '@ascend/sdk'
import { jsonOut } from '../output/json.ts'

export async function calcCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	if (!file) {
		console.error('Usage: ascend calc <file>')
		return 1
	}

	const wb = await AscendWorkbook.open(file)
	const result = wb.recalc()

	if (result.errors.length > 0) {
		for (const e of result.errors) console.error(`${e.ref}: ${e.error.message}`)
	}

	await wb.save(file)

	if (flags.has('json')) {
		console.log(jsonOut(result))
	} else {
		console.log(`Recalculated ${file} in ${result.duration}ms`)
		console.log(`  Changed: ${result.changed.length} cell(s)`)
		if (result.errors.length > 0) {
			console.log(`  Errors: ${result.errors.length}`)
		}
	}
	return 0
}
