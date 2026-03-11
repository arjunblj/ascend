#!/usr/bin/env bun

import { calcCommand, usage as calcUsage } from './commands/calc.ts'
import { checkCommand, usage as checkUsage } from './commands/check.ts'
import { createCommand, usage as createUsage } from './commands/create.ts'
import { diffCommand, usage as diffUsage } from './commands/diff.ts'
import { doctorCommand, usage as doctorUsage } from './commands/doctor.ts'
import { exportCommand, usage as exportUsage } from './commands/export.ts'
import { formulaCommand, usage as formulaUsage } from './commands/formula.ts'
import { inspectCommand, usage as inspectUsage } from './commands/inspect.ts'
import { lintCommand, usage as lintUsage } from './commands/lint.ts'
import { previewCommand, usage as previewUsage } from './commands/preview.ts'
import { readCommand, usage as readUsage } from './commands/read.ts'
import { traceCommand, usage as traceUsage } from './commands/trace.ts'
import { writeCommand, usage as writeUsage } from './commands/write.ts'
import { jsonErr } from './output/json.ts'

const VERSION = '0.0.0'

const HELP = `ascend — spreadsheet engine CLI

Usage: ascend <command> [args] [flags]

Commands:
  create <file>                 Create a new empty .xlsx workbook
  inspect <file> [sheet]        Show workbook/sheet structure
  read <file> <range>           Read cell values from a range
  preview <file> <range> <json> Preview workbook changes without saving
  write <file> <range> <json>   Write values to cells
  formula <subcommand>          Inspect or edit formulas
  calc <file>                   Recalculate all formulas
  check <file>                  Run structural checks
  lint <file>                   Run formula lint
  trace <file> <cell>           Trace precedents/dependents
  diff <file-a> <file-b>        Semantic diff between two workbooks
  export <file> <output>        Export workbook (csv, json, xlsx)
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
	inspect: {
		run: inspectCommand,
		usage: inspectUsage,
		allowedFlags: ['sheet', 'detail', 'mode', 'json', 'verbose'],
	},
	read: {
		run: readCommand,
		usage: readUsage,
		allowedFlags: ['sheet', 'mode', 'row-offset', 'row-limit', 'display', 'json'],
	},
	preview: { run: previewCommand, usage: previewUsage, allowedFlags: ['sheet', 'ops', 'json'] },
	write: { run: writeCommand, usage: writeUsage, allowedFlags: ['sheet', 'ops', 'json'] },
	formula: { run: formulaCommand, usage: formulaUsage, allowedFlags: ['json'] },
	calc: { run: calcCommand, usage: calcUsage, allowedFlags: ['json'] },
	check: { run: checkCommand, usage: checkUsage, allowedFlags: ['json'] },
	lint: { run: lintCommand, usage: lintUsage, allowedFlags: ['json'] },
	trace: { run: traceCommand, usage: traceUsage, allowedFlags: ['json', 'max-depth'] },
	diff: { run: diffCommand, usage: diffUsage, allowedFlags: ['json'] },
	export: { run: exportCommand, usage: exportUsage, allowedFlags: ['format', 'sheet', 'json'] },
	doctor: { run: doctorCommand, usage: doctorUsage },
}

const GLOBAL_FLAGS = new Set(['help', 'h', 'version', 'v'])

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
			const key = arg.slice(2)
			const next = argv[i + 1]
			if (next && !next.startsWith('-')) {
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

async function main(): Promise<void> {
	const { command, args, flags } = parseArgs(process.argv.slice(2))

	if (flags.has('version') || flags.has('v')) {
		console.log(VERSION)
		process.exit(0)
	}

	if (!command) {
		console.log(HELP)
		process.exit(0)
	}

	const cmd = COMMANDS[command]
	if (!cmd) {
		if (flags.has('help') || flags.has('h')) {
			console.log(HELP)
			process.exit(0)
		}
		console.error(`Unknown command: ${command}`)
		const suggestion = suggestClosest(command, Object.keys(COMMANDS))
		if (suggestion) console.error(`Did you mean "${suggestion}"?`)
		console.error('Run "ascend --help" for usage')
		process.exit(1)
	}

	if (flags.has('help') || flags.has('h')) {
		console.log(cmd.usage)
		process.exit(0)
	}

	const invalidFlags = [...flags.keys()].filter(
		(flag) => !GLOBAL_FLAGS.has(flag) && !(cmd.allowedFlags ?? []).includes(flag),
	)
	if (invalidFlags.length > 0) {
		const invalid = invalidFlags[0] ?? ''
		console.error(`Unknown flag for "${command}": --${invalid}`)
		const suggestion = suggestClosest(invalid, cmd.allowedFlags ?? [])
		if (suggestion) console.error(`Did you mean "--${suggestion}"?`)
		console.error(cmd.usage)
		process.exit(1)
	}

	try {
		const code = await cmd.run(args, flags)
		process.exit(code)
	} catch (err) {
		const message = err instanceof Error ? err.message : String(err)
		if (flags.has('json')) {
			console.log(jsonErr(message))
		} else {
			console.error(`Error: ${message}`)
		}
		process.exit(1)
	}
}

main()

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

function levenshtein(a: string, b: string): number {
	if (a === b) return 0
	if (a.length === 0) return b.length
	if (b.length === 0) return a.length
	const prev = Array.from({ length: b.length + 1 }, (_, i) => i)
	const curr = new Array<number>(b.length + 1).fill(0)
	for (let i = 0; i < a.length; i++) {
		curr[0] = i + 1
		for (let j = 0; j < b.length; j++) {
			const left = curr[j] ?? 0
			const up = prev[j + 1] ?? 0
			const diag = prev[j] ?? 0
			const cost = (a[i] ?? '') === (b[j] ?? '') ? 0 : 1
			curr[j + 1] = Math.min(left + 1, up + 1, diag + cost)
		}
		for (let j = 0; j < prev.length; j++) prev[j] = curr[j] ?? 0
	}
	return prev[b.length] ?? Math.max(a.length, b.length)
}
