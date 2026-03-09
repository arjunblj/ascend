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
import { readCommand, usage as readUsage } from './commands/read.ts'
import { traceCommand, usage as traceUsage } from './commands/trace.ts'
import { writeCommand, usage as writeUsage } from './commands/write.ts'

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
  --verbose     Show extended details
  --help, -h    Show help (use "ascend <command> --help" for command help)
  --version, -v Show version
`

type CommandFn = (args: string[], flags: Map<string, string>) => Promise<number>

interface Command {
	run: CommandFn
	usage: string
}

const COMMANDS: Record<string, Command> = {
	create: { run: createCommand, usage: createUsage },
	inspect: { run: inspectCommand, usage: inspectUsage },
	read: { run: readCommand, usage: readUsage },
	write: { run: writeCommand, usage: writeUsage },
	formula: { run: formulaCommand, usage: formulaUsage },
	calc: { run: calcCommand, usage: calcUsage },
	check: { run: checkCommand, usage: checkUsage },
	lint: { run: lintCommand, usage: lintUsage },
	trace: { run: traceCommand, usage: traceUsage },
	diff: { run: diffCommand, usage: diffUsage },
	export: { run: exportCommand, usage: exportUsage },
	doctor: { run: doctorCommand, usage: doctorUsage },
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
		console.error('Run "ascend --help" for usage')
		process.exit(1)
	}

	if (flags.has('help') || flags.has('h')) {
		console.log(cmd.usage)
		process.exit(0)
	}

	try {
		const code = await cmd.run(args, flags)
		process.exit(code)
	} catch (err) {
		const message = err instanceof Error ? err.message : String(err)
		console.error(`Error: ${message}`)
		process.exit(1)
	}
}

main()
