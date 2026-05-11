#!/usr/bin/env bun

import { AscendException, ascendError, levenshtein } from '@ascend/schema'
import { agentInitCommand, usage as agentInitUsage } from './commands/agent-init.ts'
import { agentViewCommand, usage as agentViewUsage } from './commands/agent-view.ts'
import { calcCommand, usage as calcUsage } from './commands/calc.ts'
import { capabilitiesCommand, usage as capabilitiesUsage } from './commands/capabilities.ts'
import { checkCommand, usage as checkUsage } from './commands/check.ts'
import { commitCommand, usage as commitUsage } from './commands/commit.ts'
import { createCommand, usage as createUsage } from './commands/create.ts'
import { diffCommand, usage as diffUsage } from './commands/diff.ts'
import { docsCommand, usage as docsUsage } from './commands/docs.ts'
import { doctorCommand, usage as doctorUsage } from './commands/doctor.ts'
import { exportCommand, usage as exportUsage } from './commands/export.ts'
import { findCommand, usage as findUsage } from './commands/find.ts'
import { formulaCommand, usage as formulaUsage } from './commands/formula.ts'
import { inspectCommand, usage as inspectUsage } from './commands/inspect.ts'
import { lintCommand, usage as lintUsage } from './commands/lint.ts'
import { listCommand, usage as listUsage } from './commands/list.ts'
import { openCommand, usage as openUsage } from './commands/open.ts'
import { opsCommand, usage as opsUsage } from './commands/ops.ts'
import { planCommand, usage as planUsage } from './commands/plan.ts'
import { previewCommand, usage as previewUsage } from './commands/preview.ts'
import { readCommand, usage as readUsage } from './commands/read.ts'
import { repairPlanCommand, usage as repairPlanUsage } from './commands/repair-plan.ts'
import { traceCommand, usage as traceUsage } from './commands/trace.ts'
import { tuiCommand, usage as tuiUsage } from './commands/tui.ts'
import { writeCommand, usage as writeUsage } from './commands/write.ts'
import { jsonErr } from './output/json.ts'

const pkg = await import('../package.json')
const VERSION = pkg.version ?? '0.0.0'

const HELP = `ascend — spreadsheet engine CLI

Usage: ascend <command> [args] [flags]

Commands:
  create <file>                 Create a new empty .xlsx workbook
  inspect <file> [sheet]        Show workbook/sheet structure
  list <file>                   List sheets and tables
  read <file> <range>           Read cell values from a range
  find <file> <query>           Search for cells matching a value
  agent-view <file>             Get AI-friendly sheet summary
  agent-init                    Print the recommended agent workflow contract
  ops                           List operation schemas and examples
  capabilities                  Show Excel capability coverage matrix
  plan <file> --ops <json>      Validate and preview a safe edit plan
  commit <file> --ops <json>    Commit an edit plan atomically
  repair-plan <file>            Suggest next actions for unsafe files
  docs [query]                  Search bundled agent docs and examples
  preview <file> <range> <json> Preview workbook changes without saving
  write <file> <range> <json>   Write values to cells
  formula <subcommand>          Inspect or edit formulas
  calc <file>                   Recalculate all formulas
  check <file>                  Run structural checks
  lint <file>                   Run formula lint
  trace <file> <cell>           Trace precedents/dependents
  diff <file-a> <file-b>        Semantic diff between two workbooks
  export <file> <output>        Export workbook (csv, json, xlsx)
  tui [file]                    Interactive terminal spreadsheet
  open [file]                   Friendly terminal spreadsheet entrypoint
  doctor                        Verify environment

Global flags:
  --help, -h    Show help (use "ascend <command> --help" for command help)
  --version, -v Show version

Common command flags:
  --json        Output structured JSON when supported
  --verbose     Show extended details on supported commands
`

type CommandFn = (args: string[], flags: Map<string, string>) => Promise<number>

interface Command {
	run: CommandFn
	usage: string
	allowedFlags?: readonly string[]
}

const COMMANDS: Record<string, Command> = {
	create: { run: createCommand, usage: createUsage },
	'agent-init': { run: agentInitCommand, usage: agentInitUsage, allowedFlags: ['json'] },
	'agent-view': {
		run: agentViewCommand,
		usage: agentViewUsage,
		allowedFlags: ['sheet', 'range', 'json'],
	},
	inspect: {
		run: inspectCommand,
		usage: inspectUsage,
		allowedFlags: ['sheet', 'detail', 'mode', 'json', 'verbose'],
	},
	list: { run: listCommand, usage: listUsage, allowedFlags: ['json'] },
	read: {
		run: readCommand,
		usage: readUsage,
		allowedFlags: ['sheet', 'mode', 'row-offset', 'row-limit', 'display', 'json'],
	},
	ops: { run: opsCommand, usage: opsUsage, allowedFlags: ['op', 'json'] },
	capabilities: {
		run: capabilitiesCommand,
		usage: capabilitiesUsage,
		allowedFlags: ['feature', 'family', 'priority', 'status', 'gaps', 'json'],
	},
	plan: { run: planCommand, usage: planUsage, allowedFlags: ['ops', 'progress', 'json'] },
	commit: {
		run: commitCommand,
		usage: commitUsage,
		allowedFlags: [
			'ops',
			'output',
			'in-place',
			'backup',
			'expect-sha256',
			'allow-loss',
			'approval',
			'progress',
			'json',
		],
	},
	'repair-plan': { run: repairPlanCommand, usage: repairPlanUsage, allowedFlags: ['json'] },
	docs: {
		run: docsCommand,
		usage: docsUsage,
		allowedFlags: ['query', 'examples', 'path', 'list', 'limit', 'tokens', 'json'],
	},
	preview: { run: previewCommand, usage: previewUsage, allowedFlags: ['sheet', 'ops', 'json'] },
	write: { run: writeCommand, usage: writeUsage, allowedFlags: ['sheet', 'ops', 'json'] },
	formula: { run: formulaCommand, usage: formulaUsage, allowedFlags: ['json'] },
	calc: { run: calcCommand, usage: calcUsage, allowedFlags: ['json'] },
	check: { run: checkCommand, usage: checkUsage, allowedFlags: ['progress', 'json'] },
	lint: { run: lintCommand, usage: lintUsage, allowedFlags: ['json'] },
	trace: { run: traceCommand, usage: traceUsage, allowedFlags: ['json', 'max-depth'] },
	diff: { run: diffCommand, usage: diffUsage, allowedFlags: ['json'] },
	export: { run: exportCommand, usage: exportUsage, allowedFlags: ['format', 'sheet', 'json'] },
	find: {
		run: findCommand,
		usage: findUsage,
		allowedFlags: ['sheet', 'match', 'json'],
	},
	tui: {
		run: tuiCommand,
		usage: tuiUsage,
		allowedFlags: ['sheet', 'renderer', 'calibrate', 'telemetry-json'],
	},
	open: {
		run: openCommand,
		usage: openUsage,
		allowedFlags: ['sheet', 'renderer', 'calibrate', 'telemetry-json'],
	},
	doctor: { run: doctorCommand, usage: doctorUsage },
}

const GLOBAL_FLAGS = new Set(['help', 'h', 'version', 'v'])
const BOOLEAN_FLAGS = new Set([
	'help',
	'h',
	'version',
	'v',
	'json',
	'verbose',
	'in-place',
	'examples',
	'list',
	'telemetry-json',
	'calibrate',
])

function parseArgs(argv: string[]): {
	command: string | undefined
	args: string[]
	flags: Map<string, string>
} {
	const flags = new Map<string, string>()
	const positional: string[] = []

	let i = 0
	while (i < argv.length) {
		const arg = argv[i] as string
		if (arg.startsWith('--')) {
			const [key = '', inlineValue] = arg.slice(2).split(/=(.*)/s, 2)
			if (inlineValue !== undefined) {
				flags.set(key, inlineValue)
				i += 1
				continue
			}
			const next = argv[i + 1]
			if (!BOOLEAN_FLAGS.has(key) && next && !next.startsWith('-')) {
				flags.set(key, next)
				i += 2
			} else {
				flags.set(key, '')
				i += 1
			}
		} else if (arg.startsWith('-') && arg.length === 2) {
			flags.set(arg.slice(1), '')
			i += 1
		} else {
			positional.push(arg)
			i += 1
		}
	}

	return {
		command: positional[0],
		args: positional.slice(1),
		flags,
	}
}

export async function runCli(argv: string[] = process.argv.slice(2)): Promise<number> {
	const { command, args, flags } = parseArgs(argv)

	if (flags.has('version') || flags.has('v')) {
		console.log(VERSION)
		return 0
	}

	if (!command) {
		console.log(HELP)
		return 0
	}

	const cmd = COMMANDS[command]
	if (!cmd) {
		if (flags.has('help') || flags.has('h')) {
			console.log(HELP)
			return 0
		}
		const suggestion = suggestClosest(command, Object.keys(COMMANDS))
		if (flags.has('json')) {
			console.log(
				jsonErr(
					ascendError('INVALID_ARGUMENT', `Unknown command: ${command}`, {
						...(suggestion ? { details: { suggestion } } : {}),
						suggestedFix: suggestion
							? `Run "ascend ${suggestion} --help".`
							: 'Run "ascend --help" for usage.',
					}),
				),
			)
			return 1
		}
		console.error(`Unknown command: ${command}`)
		if (suggestion) console.error(`Did you mean "${suggestion}"?`)
		console.error('Run "ascend --help" for usage')
		return 1
	}

	if (flags.has('help') || flags.has('h')) {
		console.log(cmd.usage)
		return 0
	}

	const invalidFlags = [...flags.keys()].filter(
		(flag) => !GLOBAL_FLAGS.has(flag) && !(cmd.allowedFlags ?? []).includes(flag),
	)
	if (invalidFlags.length > 0) {
		const invalid = invalidFlags[0] ?? ''
		const suggestion = suggestClosest(invalid, cmd.allowedFlags ?? [])
		if (flags.has('json')) {
			console.log(
				jsonErr(
					ascendError('INVALID_ARGUMENT', `Unknown flag for "${command}": --${invalid}`, {
						...(suggestion ? { details: { suggestion } } : {}),
						suggestedFix: suggestion ? `Use "--${suggestion}".` : cmd.usage,
					}),
				),
			)
			return 1
		}
		console.error(`Unknown flag for "${command}": --${invalid}`)
		if (suggestion) console.error(`Did you mean "--${suggestion}"?`)
		console.error(cmd.usage)
		return 1
	}

	try {
		return await cmd.run(args, flags)
	} catch (err) {
		if (flags.has('json')) {
			const error =
				err instanceof AscendException
					? err.ascendError
					: err instanceof Error
						? err.message
						: String(err)
			console.log(jsonErr(error))
		} else {
			console.error(`Error: ${err instanceof Error ? err.message : String(err)}`)
		}
		return 1
	}
}

if (import.meta.main) {
	const code = await runCli()
	process.exit(code)
}

function suggestClosest(input: string, candidates: readonly string[]): string | undefined {
	let best: { candidate: string; distance: number } | undefined
	for (const candidate of candidates) {
		const distance = levenshtein(input, candidate)
		if (!best || distance < best.distance) best = { candidate, distance }
	}
	if (!best) return undefined
	return best.distance <= Math.max(2, Math.floor(best.candidate.length / 3))
		? best.candidate
		: undefined
}
