import type { VbaModuleInfo, VbaModuleKind, VbaProjectInfo } from '@ascend/core'
import { isCompoundFile, readCompoundFileDirectory, readCompoundFileStream } from './cfb.ts'

const PROJECT_STREAM_NAME = 'PROJECT'
const MODULE_LINE_PATTERN = /^(Document|Module|Class|BaseClass)=(.+)$/i
const MAX_MODULE_NAME_LENGTH = 255

const moduleKindByKey: Readonly<Record<string, VbaModuleKind>> = {
	baseclass: 'designer',
	class: 'class',
	document: 'document',
	module: 'standard',
}

export function summarizeVbaProject(bytes: Uint8Array): VbaProjectInfo | undefined {
	if (!isCompoundFile(bytes)) return undefined
	try {
		const directory = readCompoundFileDirectory(bytes)
		const project = readCompoundFileStream(bytes, PROJECT_STREAM_NAME)
		if (!project) {
			return {
				moduleCount: 0,
				modules: [],
				projectStreamPresent: false,
				cfbDirectoryEntryCount: directory.length,
			}
		}
		const modules = parseProjectModules(project)
		return {
			moduleCount: modules.length,
			modules,
			projectStreamPresent: true,
			cfbDirectoryEntryCount: directory.length,
		}
	} catch {
		return undefined
	}
}

function parseProjectModules(projectStream: Uint8Array): readonly VbaModuleInfo[] {
	const modules: VbaModuleInfo[] = []
	const seen = new Set<string>()
	for (const rawLine of decodeProjectStream(projectStream).split(/\r\n|\n|\r/)) {
		const match = MODULE_LINE_PATTERN.exec(rawLine.trim())
		if (!match) continue
		const kind = moduleKindByKey[match[1]?.toLowerCase() ?? '']
		if (!kind) continue
		const name = normalizeModuleName(match[2] ?? '', kind)
		if (!name) continue
		const key = `${kind}:${name.toLowerCase()}`
		if (seen.has(key)) continue
		seen.add(key)
		modules.push({ name, kind })
	}
	return modules
}

function normalizeModuleName(value: string, kind: VbaModuleKind): string {
	const rawName = kind === 'document' ? (value.split('/')[0] ?? '') : value
	const unquoted = rawName.trim().replace(/^"|"$/g, '')
	const printable = [...unquoted]
		.filter((char) => {
			const code = char.charCodeAt(0)
			return code >= 0x20 && code !== 0x7f
		})
		.join('')
		.trim()
	return printable.slice(0, MAX_MODULE_NAME_LENGTH)
}

function decodeProjectStream(projectStream: Uint8Array): string {
	const decoder = new TextDecoder('latin1')
	return decoder.decode(projectStream)
}
