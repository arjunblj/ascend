import { AscendWorkbook } from '@ascend/sdk'
import { jsonOut } from '../output/json.ts'

export async function createCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	if (!file) {
		console.error('Usage: ascend create <file>')
		return 1
	}

	const wb = AscendWorkbook.create()
	await wb.save(file)

	if (flags.has('json')) {
		console.log(jsonOut({ created: file }))
	} else {
		console.log(`Created ${file}`)
	}
	return 0
}
