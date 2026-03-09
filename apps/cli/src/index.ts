#!/usr/bin/env bun

import { calcCommand } from './commands/calc.ts'
import { checkCommand } from './commands/check.ts'
import { createCommand } from './commands/create.ts'
import { diffCommand } from './commands/diff.ts'
import { doctorCommand } from './commands/doctor.ts'
import { exportCommand } from './commands/export.ts'
import { formulaCommand } from './commands/formula.ts'
import { inspectCommand } from './commands/inspect.ts'
import { lintCommand } from './commands/lint.ts'
import { readCommand } from './commands/read.ts'
import { traceCommand } from './commands/trace.ts'
import { writeCommand } from './commands/write.ts'

const VERSION = '0.0.0'

const HELP = `ascend — spreadsheet engine CLI

Usage: ascend <command> [args] [flags]

Commands:
  create <file>                 Create a new empty .xlsx workbook
  inspect <file> [sheet]        Show workbook/sheet structure
  read <file> <range>           Read cell values from a range
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
  --json        Output as JSON
  --help, -h    Show help
  --version, -v Show version
`

type CommandFn = (args: string[], flags: Map<string, string>) => Promise<number>

const COMMANDS: Record<string, CommandFn> = {
	create: createCommand,
	inspect: inspectCommand,
	read: readCommand,
	write: writeCommand,
	formula: formulaCommand,
	calc: calcCommand,
	check: checkCommand,
	lint: lintCommand,
	trace: traceCommand,
	diff: diffCommand,
	export: exportCommand,
	doctor: doctorCommand,
}

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

	if (!command || flags.has('help') || flags.has('h')) {
		console.log(HELP)
		process.exit(0)
	}

	const handler = COMMANDS[command]
	if (!handler) {
		console.error(`Unknown command: ${command}`)
		console.error('Run "ascend --help" for usage')
		process.exit(1)
	}

	try {
		const code = await handler(args, flags)
		process.exit(code)
	} catch (err) {
		const message = err instanceof Error ? err.message : String(err)
		console.error(`Error: ${message}`)
		process.exit(1)
	}
}

main()
