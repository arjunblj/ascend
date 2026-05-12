import { cliError, jsonOut } from '../output/json.ts'
import { table } from '../output/pretty.ts'
import {
	createJsonlProgressWriter,
	emitCliProgress,
	openWorkbookDocumentWithProgress,
} from '../progress.ts'

export const usage = `Usage: ascend check <file> [flags]

  Run structural checks on a workbook.

Arguments:
  <file>          Path to the workbook file

Flags:
  --progress jsonl  Emit machine-readable progress events to stderr
  --json          Output as JSON
`

export async function checkCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	if (!file) {
		cliError('Usage: ascend check <file>', flags)
		return 1
	}

	const progress = createJsonlProgressWriter(flags)
	emitCliProgress(progress, {
		sequence: 1,
		kind: 'check',
		phase: 'load-workbook',
		status: 'started',
		summary: 'Opening workbook for structural checks.',
	})
	const { document: wb } = await openWorkbookDocumentWithProgress(file)
	emitCliProgress(progress, {
		sequence: 2,
		kind: 'check',
		phase: 'load-workbook',
		status: 'ok',
		summary: 'Workbook opened.',
	})
	emitCliProgress(progress, {
		sequence: 3,
		kind: 'check',
		phase: 'check',
		status: 'started',
		summary: 'Running structural checks.',
	})
	const result = wb.check()
	emitCliProgress(progress, {
		sequence: 4,
		kind: 'check',
		phase: 'check',
		status: result.valid ? 'ok' : 'failed',
		summary: result.valid ? 'Structural checks passed.' : `${result.issues.length} issue(s) found.`,
		count: result.issues.length,
	})

	if (flags.has('json')) {
		console.log(jsonOut(result))
	} else if (result.valid) {
		console.log(`${file}: all checks passed`)
	} else {
		console.log(`${file}: ${result.issues.length} issue(s) found\n`)
		console.log(
			table(
				['Severity', 'Rule', 'Kind', 'Message', 'Ref', 'Suggested Fix'],
				result.issues.map((i) => [
					i.severity,
					i.rule ?? '',
					issueKind(i.details),
					i.message,
					i.ref ?? '',
					i.suggestedFix ?? '',
				]),
			),
		)
	}

	return result.valid ? 0 : 2
}

function issueKind(details: unknown): string {
	if (!details || typeof details !== 'object') return ''
	const kind = (details as { readonly kind?: unknown }).kind
	return typeof kind === 'string' ? kind : ''
}
