export const usage = `Usage: ascend tui [file] [flags]

  Start the optional interactive terminal spreadsheet.

Flags:
  --sheet <name>         Open a specific sheet
  --preview-rows <n>    Limit initial preview rows
  --renderer <name>     Renderer selection
  --calibrate           Show terminal calibration
  --telemetry-json      Emit telemetry JSON
`

type TuiModule = {
	tuiCommand: (args: string[], flags: Map<string, string>) => Promise<number>
}

export async function tuiCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const mod = await loadTui()
	if (!mod) return 1
	return mod.tuiCommand(args, flags)
}

async function loadTui(): Promise<TuiModule | null> {
	try {
		return (await import('@ascend/tui')) as TuiModule
	} catch (error) {
		console.error(
			'The interactive TUI is not installed in this headless Ascend CLI package. Install @ascend/tui or use headless commands like inspect, plan, commit, and check.',
		)
		if (!(error instanceof Error) || !error.message.includes('@ascend/tui')) {
			console.error(error instanceof Error ? error.message : String(error))
		}
		return null
	}
}
