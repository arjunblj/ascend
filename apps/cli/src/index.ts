#!/usr/bin/env bun

import { pathToFileURL } from 'node:url'
import { type AscendError, AscendException, ascendError, levenshtein } from '@ascend/schema'
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
import { dumpCommand, usage as dumpUsage } from './commands/dump.ts'
import {
	exampleSafeEditCommand,
	usage as exampleSafeEditUsage,
} from './commands/example-safe-edit.ts'
import { exportCommand, usage as exportUsage } from './commands/export.ts'
import { findCommand, usage as findUsage } from './commands/find.ts'
import { formulaCommand, usage as formulaUsage } from './commands/formula.ts'
import { inspectCommand, usage as inspectUsage } from './commands/inspect.ts'
import { lintCommand, usage as lintUsage } from './commands/lint.ts'
import { listCommand, usage as listUsage } from './commands/list.ts'
import { openCommand, usage as openUsage } from './commands/open.ts'
import { openPlanCommand, usage as openPlanUsage } from './commands/open-plan.ts'
import { opsCommand, usage as opsUsage } from './commands/ops.ts'
import { planCommand, usage as planUsage } from './commands/plan.ts'
import { previewCommand, usage as previewUsage } from './commands/preview.ts'
import { readCommand, usage as readUsage } from './commands/read.ts'
import { repairPlanCommand, usage as repairPlanUsage } from './commands/repair-plan.ts'
import { templateMergeCommand, usage as templateMergeUsage } from './commands/template-merge.ts'
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
  read <file> <selector>        Read cells from a range, table, or name
  dump <file>                   Dump a replayable operation batch
  template-merge <file>         Compile template placeholders into operations
  find <file> <query>           Search for cells matching a value
  agent-view <file>             Get AI-friendly sheet summary
  agent-init                    Print the recommended agent workflow contract
  example-safe-edit <file>      Run packaged safe-edit workflow example
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
  open-plan <file>              Recommend an open mode before hydration
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
	create: { run: createCommand, usage: createUsage, allowedFlags: ['json'] },
	'agent-init': { run: agentInitCommand, usage: agentInitUsage, allowedFlags: ['json'] },
	'example-safe-edit': {
		run: exampleSafeEditCommand,
		usage: exampleSafeEditUsage,
		allowedFlags: ['json'],
	},
	'agent-view': {
		run: agentViewCommand,
		usage: agentViewUsage,
		allowedFlags: ['sheet', 'range', 'tokens', 'json'],
	},
	inspect: {
		run: inspectCommand,
		usage: inspectUsage,
		allowedFlags: ['sheet', 'detail', 'mode', 'agent', 'json', 'verbose'],
	},
	list: { run: listCommand, usage: listUsage, allowedFlags: ['json'] },
	read: {
		run: readCommand,
		usage: readUsage,
		allowedFlags: ['sheet', 'mode', 'row-offset', 'row-limit', 'display', 'json'],
	},
	dump: {
		run: dumpCommand,
		usage: dumpUsage,
		allowedFlags: ['sheet', 'values-only', 'formulas-only', 'json'],
	},
	'template-merge': {
		run: templateMergeCommand,
		usage: templateMergeUsage,
		allowedFlags: ['data', 'sheet', 'values-only', 'formulas-only', 'open', 'close', 'json'],
	},
	ops: { run: opsCommand, usage: opsUsage, allowedFlags: ['op', 'json'] },
	capabilities: {
		run: capabilitiesCommand,
		usage: capabilitiesUsage,
		allowedFlags: ['feature', 'family', 'priority', 'status', 'gaps', 'json'],
	},
	plan: {
		run: planCommand,
		usage: planUsage,
		allowedFlags: ['ops', 'password', 'package-actions', 'progress', 'json'],
	},
	commit: {
		run: commitCommand,
		usage: commitUsage,
		allowedFlags: [
			'ops',
			'output',
			'in-place',
			'backup',
			'password',
			'expect-sha256',
			'allow-loss',
			'approval',
			'progress',
			'compact',
			'package-actions',
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
	formula: {
		run: formulaCommand,
		usage: formulaUsage,
		allowedFlags: [
			'cursor',
			'prefix',
			'completion-limit',
			'function-name',
			'reference',
			'replace-reference-at-cursor',
			'cycle-reference',
			'json',
		],
	},
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
		allowedFlags: ['sheet', 'preview-rows', 'renderer', 'calibrate', 'telemetry-json'],
	},
	open: {
		run: openCommand,
		usage: openUsage,
		allowedFlags: ['sheet', 'preview-rows', 'renderer', 'calibrate', 'telemetry-json'],
	},
	'open-plan': {
		run: openPlanCommand,
		usage: openPlanUsage,
		allowedFlags: ['intent', 'password', 'json'],
	},
	doctor: { run: doctorCommand, usage: doctorUsage },
}

const GLOBAL_FLAGS = new Set(['help', 'h', 'version', 'v'])
const AGENT_WORKFLOW = ['inspect', 'plan', 'commit', 'reopen', 'verify'] as const
const BOOLEAN_FLAGS = new Set([
	'help',
	'h',
	'version',
	'v',
	'json',
	'verbose',
	'agent',
	'in-place',
	'examples',
	'list',
	'compact',
	'package-actions',
	'telemetry-json',
	'calibrate',
	'replace-reference-at-cursor',
	'cycle-reference',
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
						retryable: true,
						retryStrategy: 'modified',
						details: {
							command,
							availableCommands: Object.keys(COMMANDS).sort(),
							workflow: AGENT_WORKFLOW,
							...(suggestion ? { suggestion } : {}),
						},
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
						retryable: true,
						retryStrategy: 'modified',
						details: {
							command,
							flag: invalid,
							allowedFlags: [...(cmd.allowedFlags ?? [])].sort(),
							globalFlags: [...GLOBAL_FLAGS].sort(),
							workflow: AGENT_WORKFLOW,
							...(suggestion ? { suggestion } : {}),
						},
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
					: isFileNotFoundError(err)
						? fileNotFoundCliError(missingFilePath(err) ?? args[0])
						: err instanceof Error
							? err.message
							: String(err)
			console.log(jsonErr(error))
		} else {
			if (isFileNotFoundError(err)) {
				const error = fileNotFoundCliError(missingFilePath(err) ?? args[0])
				console.error(`Error: ${error.message}`)
				if (error.suggestedFix) console.error(error.suggestedFix)
			} else {
				console.error(`Error: ${err instanceof Error ? err.message : String(err)}`)
			}
		}
		return 1
	}
}

if (isDirectRun()) {
	const code = await runCli()
	process.exit(code)
}

function isDirectRun(): boolean {
	return process.argv[1] ? import.meta.url === pathToFileURL(process.argv[1]).href : false
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

function isFileNotFoundError(e: unknown): boolean {
	if (!e || typeof e !== 'object') return false
	if ((e as { readonly code?: unknown }).code === 'ENOENT') return true
	if (e instanceof Error && e.message.includes('ENOENT: no such file or directory')) return true
	return false
}

function missingFilePath(e: unknown): string | undefined {
	if (!e || typeof e !== 'object') return undefined
	const path = (e as { readonly path?: unknown }).path
	return typeof path === 'string' && path.length > 0 ? path : undefined
}

function fileNotFoundCliError(file?: string): AscendError {
	return ascendError('FILE_NOT_FOUND', file ? `File not found: ${file}` : 'File not found', {
		retryable: true,
		retryStrategy: 'modified',
		...(file ? { details: { file } } : {}),
		suggestedFix:
			'Pass an existing workbook path that this CLI process can read, then retry the command.',
	})
}
