import { createServer } from './server.ts'

const server = createServer()

console.log(`Server running at http://localhost:${server.port}`)
