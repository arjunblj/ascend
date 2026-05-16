import { readFileSync } from 'node:fs'
import { ascendError } from '@ascend/schema'
import {
	inspectWorkbookOpenPlan,
	type WorkbookOpenIntent,
	type WorkbookOpenPlan,
} from '@ascend/sdk'
import { cliError, jsonOut } from '../output/json.ts'
import { bullet, heading, table } from '../output/pretty.ts'

export const usage = `Usage: ascend open-plan <file> [flags]

  Recommend a safe XLSX open mode from package features before workbook hydration.

Arguments:
  <file>              Path to the workbook file

Flags:
  --intent <intent>   risk-inventory, read-values, formula-analysis, or edit-plan (default: edit-plan)
  --password <value>  Password for encrypted XLSX/XLSM workbooks
  --json              Output as JSON
`

export async function openPlanCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	if (!file) {
		cliError(
			ascendError('INVALID_ARGUMENT', 'Missing required open-plan input', {
				retryable: true,
				retryStrategy: 'modified',
				details: {
					command: 'open-plan',
					required: ['file'],
					missing: ['file'],
					workflow: ['open-plan', 'inspect', 'plan', 'commit', 'reopen', 'verify'],
				},
				suggestedFix:
					'Run ascend open-plan <file> --json before hydrating an unknown XLSX/XLSM workbook.',
			}),
			flags,
		)
		return 1
	}

	const intent = parseOpenIntent(flags.get('intent'))
	if (flags.has('intent') && intent === null) {
		cliError(
			'Invalid --intent. Use one of: risk-inventory, read-values, formula-analysis, edit-plan',
			flags,
		)
		return 1
	}

	const password = flags.get('password')
	const plan = inspectWorkbookOpenPlan(new Uint8Array(readFileSync(file)), {
		...(intent ? { intent } : {}),
		...(password !== undefined ? { password } : {}),
	})
	if (flags.has('json')) {
		console.log(jsonOut(plan))
		return 0
	}

	printOpenPlan(file, plan)
	return 0
}

function parseOpenIntent(intent: string | undefined): WorkbookOpenIntent | undefined | null {
	if (intent === undefined || intent === '') return undefined
	switch (intent) {
		case 'risk-inventory':
		case 'read-values':
		case 'formula-analysis':
		case 'edit-plan':
			return intent
		default:
			return null
	}
}

function printOpenPlan(file: string, plan: WorkbookOpenPlan): void {
	console.log(heading(`Open plan: ${file}`))
	console.log(bullet('Intent', plan.intent))
	console.log(bullet('Recommended mode', plan.recommendedMode))
	console.log(bullet('Recommended load options', JSON.stringify(plan.recommendedLoadOptions)))
	console.log(bullet('Rich metadata', plan.richMetadataRecommended ? 'yes' : 'no'))
	console.log(bullet('Review before hydration', plan.reviewBeforeHydration ? 'yes' : 'no'))
	console.log(bullet('Cost class', plan.costClass))
	console.log(bullet('Parts', plan.partCount))
	console.log(bullet('Worksheet parts', plan.worksheetPartCount))
	console.log(bullet('Relationships', plan.relationshipCount))
	console.log(bullet('Formula signal', plan.formulaSignal ? 'yes' : 'no'))
	if (plan.reasons.length > 0) {
		console.log(heading('Reasons'))
		for (const reason of plan.reasons) console.log(`  - ${reason}`)
	}
	if (plan.riskFeatures.length > 0) {
		console.log(heading('Risk Features'))
		console.log(
			table(
				['Feature', 'Category', 'Count', 'Sample parts'],
				plan.riskFeatures.map((feature) => [
					feature.featureFamily,
					feature.category,
					String(feature.count),
					feature.sampleParts.join(', '),
				]),
			),
		)
	}
}
