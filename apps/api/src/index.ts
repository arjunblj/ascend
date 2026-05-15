import { pathToFileURL } from 'node:url'

export { createApiFetch, createServer } from './server.ts'

if (isDirectRun()) {
	const { createServer } = await import('./server.ts')
	const server = createServer()

	console.log(`Server running at http://localhost:${server.port}`)
}

function isDirectRun(): boolean {
	return process.argv[1] ? import.meta.url === pathToFileURL(process.argv[1]).href : false
}
