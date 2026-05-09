import { mkdirSync, readdirSync, writeFileSync } from 'node:fs'
import { basename } from 'node:path'

type ProfileMode = 'cpu' | 'cpu-md' | 'heap' | 'heap-md' | 'all' | 'all-md'

interface Config {
	readonly mode: ProfileMode
	readonly label: string
	readonly outDir: string
	readonly command: readonly string[]
	readonly requireOutput: boolean
}

const MODES = new Set<ProfileMode>(['cpu', 'cpu-md', 'heap', 'heap-md', 'all', 'all-md'])

function main(): Promise<void> {
	const config = parseArgs(process.argv.slice(2))
	mkdirSync(config.outDir, { recursive: true })

	const stamp = new Date().toISOString().replace(/[:.]/g, '-')
	const baseName = `${config.label}-${stamp}`
	const args = [
		...profileFlags(config.mode, baseName, config.outDir),
		...normalizeCommand(config.command),
	]

	console.error(`bun profile: ${config.mode}`)
	console.error(`profile dir: ${config.outDir}`)
	console.error(`command: ${process.execPath} ${args.join(' ')}`)

	const child = Bun.spawn([process.execPath, ...args], {
		cwd: process.cwd(),
		env: process.env,
		stdout: 'inherit',
		stderr: 'inherit',
	})
	return child.exited.then((exitCode) => {
		const files = readdirSync(config.outDir).filter((file) => file.startsWith(baseName))
		if (files.length > 0) {
			console.error(`profile files: ${files.map((file) => `${config.outDir}/${file}`).join(', ')}`)
			process.exitCode = exitCode
			return
		}
		const diagnosticPath = `${config.outDir}/${baseName}.profile-missing.txt`
		writeFileSync(diagnosticPath, missingProfileDiagnostic(config, baseName, args, exitCode))
		console.error(`profile files: none found for ${config.outDir}/${baseName}*`)
		console.error(`profile diagnostic: ${diagnosticPath}`)
		process.exitCode = exitCode === 0 && config.requireOutput ? 1 : exitCode
	})
}

function missingProfileDiagnostic(
	config: Config,
	baseName: string,
	args: readonly string[],
	exitCode: number,
): string {
	return [
		`label: ${config.label}`,
		`mode: ${config.mode}`,
		`baseName: ${baseName}`,
		`bun: ${Bun.version}+${Bun.revision}`,
		`exitCode: ${exitCode}`,
		`expectedPrefix: ${config.outDir}/${baseName}`,
		`command: ${process.execPath} ${args.join(' ')}`,
		'',
		'Bun completed the target command but did not emit a profile artifact matching the requested prefix.',
		'This can happen when the installed Bun build ignores or lacks the selected profiling flags.',
		'Use --require-output to make missing profile artifacts fail the wrapper.',
		'',
	].join('\n')
}

function profileFlags(mode: ProfileMode, baseName: string, outDir: string): string[] {
	const flags: string[] = []
	if (mode === 'cpu' || mode === 'all') {
		flags.push('--cpu-prof', `--cpu-prof-dir=${outDir}`, `--cpu-prof-name=${baseName}.cpuprofile`)
	}
	if (mode === 'cpu-md' || mode === 'all-md') {
		flags.push('--cpu-prof-md', `--cpu-prof-dir=${outDir}`, `--cpu-prof-name=${baseName}.cpu.md`)
	}
	if (mode === 'heap' || mode === 'all') {
		flags.push(
			'--heap-prof',
			`--heap-prof-dir=${outDir}`,
			`--heap-prof-name=${baseName}.heapsnapshot`,
		)
	}
	if (mode === 'heap-md' || mode === 'all-md') {
		flags.push(
			'--heap-prof-md',
			`--heap-prof-dir=${outDir}`,
			`--heap-prof-name=${baseName}.heap.md`,
		)
	}
	return flags
}

function normalizeCommand(command: readonly string[]): readonly string[] {
	const [first, ...rest] = command
	if (!first) return command
	const name = basename(first)
	return name === 'bun' ? rest : command
}

function parseArgs(argv: readonly string[]): Config {
	let mode: ProfileMode = 'cpu-md'
	let label = 'ascend-bench'
	let outDir = 'profiles/bun'
	let requireOutput = false
	let commandStart = argv.indexOf('--')
	if (commandStart === -1) commandStart = argv.length

	for (let i = 0; i < commandStart; i++) {
		const arg = argv[i]
		const next = argv[i + 1]
		if (arg === '--mode' && next) {
			if (!MODES.has(next as ProfileMode)) {
				throw new Error(`Unsupported --mode "${next}". Expected one of: ${[...MODES].join(', ')}`)
			}
			mode = next as ProfileMode
			i++
			continue
		}
		if (arg === '--label' && next) {
			label = sanitizeLabel(next)
			i++
			continue
		}
		if (arg === '--out-dir' && next) {
			outDir = next
			i++
			continue
		}
		if (arg === '--require-output') {
			requireOutput = true
			continue
		}
		if (arg === '--help' || arg === '-h') {
			printUsage()
			process.exit(0)
		}
		throw new Error(`Unsupported argument: ${arg}`)
	}

	const command = argv.slice(commandStart + 1)
	if (command.length === 0) {
		printUsage()
		throw new Error('Missing command after --')
	}
	return { mode, label, outDir, command, requireOutput }
}

function sanitizeLabel(value: string): string {
	const sanitized = value.replace(/[^a-zA-Z0-9._-]+/g, '-').replace(/^-+|-+$/g, '')
	return sanitized || 'ascend-bench'
}

function printUsage(): void {
	console.error(`Usage:
  bun run fixtures/benchmarks/profile-bun.ts [--mode cpu-md] [--label name] [--out-dir profiles/bun] -- run <script> [args...]

Modes:
  cpu      Chrome/VS Code .cpuprofile
  cpu-md   Markdown CPU profile for CLI/LLM review
  heap     Chrome DevTools .heapsnapshot on exit
  heap-md  Markdown heap profile on exit
  all      CPU JSON + heap snapshot
  all-md   CPU markdown + heap markdown

Options:
  --require-output  fail when Bun exits successfully but emits no profile artifact

Examples:
  bun run fixtures/benchmarks/profile-bun.ts --label reader-fastexcel -- run fixtures/benchmarks/xlsx-read-phase.ts --profile fastexcel-reader-65536 --phase direct --repeat 3 --warmup 1 --json
  bun run fixtures/benchmarks/profile-bun.ts --mode all-md --label writer-excelize -- run fixtures/benchmarks/xlsx-write-phase.ts --profile excelize-generation-102400x50-plain-text --repeat 1 --warmup 0 --json`)
}

await main()
