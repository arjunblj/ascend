import { AscendWorkbook } from '@ascend/sdk'
import { jsonOut } from '../output/json.ts'
import { table } from '../output/pretty.ts'

export async function lintCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	if (!file) {
		console.error('Usage: ascend lint <file>')
		return 1
	}

	const wb = await AscendWorkbook.open(file)
	const result = wb.lint()

	if (flags.has('json')) {
		console.log(jsonOut(result))
	} else if (result.clean) {
		console.log(`${file}: no lint warnings`)
	} else {
		console.log(`${file}: ${result.warnings.length} warning(s)\n`)
		console.log(
			table(
				['Rule', 'Message', 'Ref'],
				result.warnings.map((w) => [w.rule, w.message, w.ref ?? '']),
			),
		)
	}

	return result.clean ? 0 : 2
}
