import { describe, expect, test } from 'bun:test'
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { resolve } from 'node:path'
import { strToU8, zipSync } from 'fflate'
import {
	type CorpusManifestEntry,
	normalizeManifest,
	validateManifestProvenance,
} from '../corpus/manifest.ts'
import {
	benchmarkProvenanceDimensions,
	coalesceRepeatCorrectnessStatus,
	evaluateAssertions,
	extractWorkbookFeatureSummary,
	FULL_CORPUS_TARGETS,
	libraryAllowed,
	loadCorpusManifestEntries,
	normalizeAssertions,
	normalizeExternalRunnerManifestSet,
	normalizeExternalRunnerSpecs,
	normalizeExternalSampleAssertions,
	normalizeExternalSamples,
	parseLibraryAllowlist,
	QUICK_TARGETS,
	resolveExternalRunnerCommand,
	selectCorpusTargets,
	type WorkbookFeatureSummary,
	type WorkbookPackageFingerprint,
	type WorkbookShapeSummary,
} from './competitive-real-workbook.ts'

describe('real workbook library allowlist', () => {
	test('parses exact runner names', () => {
		const allowlist = parseLibraryAllowlist('ascend, sheetjs ,openpyxl')
		expect(libraryAllowed('ascend', allowlist)).toBe(true)
		expect(libraryAllowed('sheetjs', allowlist)).toBe(true)
		expect(libraryAllowed('exceljs', allowlist)).toBe(false)
		expect(parseLibraryAllowlist(' , ')).toBeUndefined()
	})
})

describe('evaluateAssertions', () => {
	test('read assertions require semantic cell reference hashes', () => {
		const expected = shape()
		const result = evaluateAssertions('read', expected, {
			...passingReadAssertions(expected),
			semanticCellRefsHash: 'wrong',
		})
		expect(result.status).toBe('semantic-mismatch')
		expect(result.assertions.semanticCellRefsHashMatches).toBe(false)
	})

	test('read assertions require semantic cell value hashes', () => {
		const expected = shape()
		const result = evaluateAssertions('read', expected, {
			...passingReadAssertions(expected),
			semanticCellValuesHash: 'wrong',
		})
		expect(result.status).toBe('semantic-mismatch')
		expect(result.assertions.semanticCellValuesHashMatches).toBe(false)
	})

	test('read assertions accept ordered hashes without sorted semantic hashes', () => {
		const expected = shape({
			orderedSemanticCellRefsHash: 'ordered-refs',
			orderedSemanticCellValuesHash: 'ordered-values',
			orderedFormulaTextHash: 'ordered-formulas',
		})
		const result = evaluateAssertions('read', expected, {
			...passingReadAssertions(expected),
			semanticCellRefsHash: 'missing-sorted-refs',
			semanticCellValuesHash: 'missing-sorted-values',
			formulaTextHash: 'missing-sorted-formulas',
			orderedSemanticCellRefsHash: 'ordered-refs',
			orderedSemanticCellValuesHash: 'ordered-values',
			orderedFormulaTextHash: 'ordered-formulas',
		})
		expect(result.status).toBe('pass')
		expect(result.assertions.semanticCellRefsHashMatches).toBe(true)
		expect(result.assertions.semanticCellValuesHashMatches).toBe(true)
		expect(result.assertions.formulaTextHashMatches).toBe(true)
		expect(result.assertions.orderedSemanticCellRefsHashMatches).toBe(true)
		expect(result.assertions.orderedSemanticCellValuesHashMatches).toBe(true)
		expect(result.assertions.orderedFormulaTextHashMatches).toBe(true)
	})

	test('read-values profile does not require formula text preservation', () => {
		const expected = shape({ formulaTextHash: 'formula-hash' })
		const result = evaluateAssertions(
			'read',
			expected,
			{
				...passingReadAssertions(expected),
				formulaCount: 0,
				formulaTextHash: 'empty-formulas',
			},
			'read-values',
		)
		expect(result.status).toBe('pass')
		expect(result.assertions.formulaPreservationRequired).toBe(false)
		expect(result.assertions.formulaTextHashMatches).toBe(true)
	})

	test('read-values profile does not require package fingerprint assertions', () => {
		const expected = shape({ packageFingerprint: packageFingerprint() })
		const result = evaluateAssertions(
			'read',
			expected,
			{
				...passingReadAssertions(expected),
				formulaCount: 0,
				formulaTextHash: 'empty-formulas',
			},
			'read-values',
		)
		expect(result.status).toBe('pass')
		expect(result.assertions.packageFingerprintRequired).toBeUndefined()
	})

	test('read formula hash is report-only when reader declares unsupported compatibility', () => {
		const expected = shape({ formulaTextHash: 'formula-hash' })
		const result = evaluateAssertions('read', expected, {
			...passingReadAssertions(expected),
			compatibility: 'has-unsupported',
			formulaTextHash: 'different',
		})
		expect(result.status).toBe('pass')
		expect(result.assertions.formulaTextHashRequired).toBe(false)
		expect(result.assertions.formulaTextHashMatches).toBe(true)
	})

	test('semantic roundtrip requires formula hashes when package bytes differ', () => {
		const expected = shape({ formulaTextHash: 'formula-hash' })
		const result = evaluateAssertions('roundtrip', expected, {
			byteIdentical: false,
			roundtripSheetCount: expected.sheetCount,
			roundtripSheetNamesHash: expected.sheetNamesHash,
			roundtripCellCount: expected.cellCount,
			roundtripFormulaCount: expected.formulaCount,
			roundtripFirstUsedRange: expected.usedRanges[0] ?? null,
			roundtripUsedRangesHash: expected.usedRangesHash,
			roundtripSemanticCellRefsHash: expected.semanticCellRefsHash,
			roundtripSemanticCellValuesHash: expected.semanticCellValuesHash,
			roundtripFormulaTextHash: 'wrong',
		})
		expect(result.status).toBe('semantic-roundtrip-mismatch')
		expect(result.assertions.roundtripFormulaTextHashMatches).toBe(false)
	})

	test('semantic roundtrip requires value hashes when package bytes differ', () => {
		const expected = shape({ semanticCellValuesHash: 'value-hash' })
		const result = evaluateAssertions('roundtrip', expected, {
			byteIdentical: false,
			roundtripSheetCount: expected.sheetCount,
			roundtripSheetNamesHash: expected.sheetNamesHash,
			roundtripCellCount: expected.cellCount,
			roundtripFormulaCount: expected.formulaCount,
			roundtripFirstUsedRange: expected.usedRanges[0] ?? null,
			roundtripUsedRangesHash: expected.usedRangesHash,
			roundtripSemanticCellRefsHash: expected.semanticCellRefsHash,
			roundtripSemanticCellValuesHash: 'wrong',
			roundtripFormulaTextHash: expected.formulaTextHash,
		})
		expect(result.status).toBe('semantic-roundtrip-mismatch')
		expect(result.assertions.roundtripSemanticCellValuesHashMatches).toBe(false)
	})

	test('semantic roundtrip requires matching package fingerprints when package bytes differ', () => {
		const expected = shape({ packageFingerprint: packageFingerprint() })
		const result = evaluateAssertions('roundtrip', expected, {
			...passingRoundtripAssertions(expected),
			byteIdentical: false,
			roundtripPackagePartNamesHash: 'wrong-package-parts',
		})
		expect(result.status).toBe('package-roundtrip-mismatch')
		expect(result.assertions.packageFingerprintRequired).toBe(true)
		expect(result.assertions.roundtripPackagePartNamesHashMatches).toBe(false)
		expect(result.assertions.semanticRoundtripMatches).toBe(true)
		expect(result.assertions.packageRoundtripMatches).toBe(false)
	})

	test('semantic roundtrip without package fingerprints is not ranking eligible', () => {
		const expected = shape({ packageFingerprint: packageFingerprint() })
		const result = evaluateAssertions('roundtrip', expected, {
			byteIdentical: false,
			roundtripSheetCount: expected.sheetCount,
			roundtripSheetNamesHash: expected.sheetNamesHash,
			roundtripCellCount: expected.cellCount,
			roundtripFormulaCount: expected.formulaCount,
			roundtripFirstUsedRange: expected.usedRanges[0] ?? null,
			roundtripUsedRangesHash: expected.usedRangesHash,
			roundtripSemanticCellRefsHash: expected.semanticCellRefsHash,
			roundtripSemanticCellValuesHash: expected.semanticCellValuesHash,
			roundtripFormulaTextHash: expected.formulaTextHash,
		})
		expect(result.status).toBe('package-roundtrip-unverified')
		expect(result.assertions.packageFingerprintRequired).toBe(true)
		expect(result.assertions.hasRoundtripPackageFingerprint).toBe(false)
		expect(result.assertions.semanticRoundtripMatches).toBe(true)
	})

	test('semantic roundtrip with matching package fingerprints is ranking eligible', () => {
		const expected = shape({ packageFingerprint: packageFingerprint() })
		const result = evaluateAssertions('roundtrip', expected, {
			...passingRoundtripAssertions(expected),
			byteIdentical: false,
		})
		expect(result.status).toBe('semantic-roundtrip-pass')
		expect(result.assertions.hasRoundtripPackageFingerprint).toBe(true)
		expect(result.assertions.packageRoundtripMatches).toBe(true)
	})

	test('semantic roundtrip without feature fingerprints is not ranking eligible', () => {
		const expected = shape({
			packageFingerprint: packageFingerprint(),
			featureSummary: featureSummary(),
		})
		const result = evaluateAssertions('roundtrip', expected, {
			...passingRoundtripAssertions(shape({ packageFingerprint: expected.packageFingerprint })),
			byteIdentical: false,
		})
		expect(result.status).toBe('feature-roundtrip-unverified')
		expect(result.assertions.featureFingerprintRequired).toBe(true)
		expect(result.assertions.hasRoundtripFeatureFingerprint).toBe(false)
		expect(result.assertions.semanticRoundtripMatches).toBe(true)
		expect(result.assertions.packageRoundtripMatches).toBe(true)
	})

	test('semantic roundtrip requires matching feature fingerprints when package bytes differ', () => {
		const expected = shape({
			packageFingerprint: packageFingerprint(),
			featureSummary: featureSummary(),
		})
		const result = evaluateAssertions('roundtrip', expected, {
			...passingRoundtripAssertions(expected),
			byteIdentical: false,
			roundtripFeatureInventoryHash: 'wrong-feature-inventory',
		})
		expect(result.status).toBe('feature-roundtrip-mismatch')
		expect(result.assertions.hasRoundtripFeatureFingerprint).toBe(true)
		expect(result.assertions.packageRoundtripMatches).toBe(true)
		expect(result.assertions.roundtripFeatureInventoryHashMatches).toBe(false)
		expect(result.assertions.featureRoundtripMatches).toBe(false)
	})

	test('semantic roundtrip with matching package and feature fingerprints is ranking eligible', () => {
		const expected = shape({
			packageFingerprint: packageFingerprint(),
			featureSummary: featureSummary(),
		})
		const result = evaluateAssertions('roundtrip', expected, {
			...passingRoundtripAssertions(expected),
			byteIdentical: false,
		})
		expect(result.status).toBe('semantic-roundtrip-pass')
		expect(result.assertions.hasRoundtripPackageFingerprint).toBe(true)
		expect(result.assertions.hasRoundtripFeatureFingerprint).toBe(true)
		expect(result.assertions.packageRoundtripMatches).toBe(true)
		expect(result.assertions.featureRoundtripMatches).toBe(true)
	})

	test('edit roundtrip allows the edited value hash to change but requires the edit to persist', () => {
		const expected = shape({
			packageFingerprint: packageFingerprint(),
			featureSummary: featureSummary(),
		})
		const result = evaluateAssertions('edit-roundtrip', expected, {
			...passingRoundtripAssertions(expected),
			roundtripSemanticCellValuesHash: 'changed-by-edit',
			editCellValueMatches: true,
		})
		expect(result.status).toBe('semantic-roundtrip-pass')
		expect(result.assertions.roundtripSemanticCellValuesHashMatches).toBe(true)
		expect(result.assertions.editCellValueMatches).toBe(true)
	})

	test('edit roundtrip fails when the edited cell value is not preserved', () => {
		const expected = shape({ packageFingerprint: packageFingerprint() })
		const result = evaluateAssertions('edit-roundtrip', expected, {
			...passingRoundtripAssertions(expected),
			roundtripSemanticCellValuesHash: 'changed-by-edit',
			editCellValueMatches: false,
		})
		expect(result.status).toBe('semantic-roundtrip-mismatch')
		expect(result.assertions.editCellValueMatches).toBe(false)
	})

	test('repeat correctness must be stable before ranking', () => {
		expect(coalesceRepeatCorrectnessStatus(['pass', 'pass', 'pass'])).toBe('pass')
		expect(coalesceRepeatCorrectnessStatus(['pass', 'semantic-mismatch', 'pass'])).toBe(
			'intermittent-mismatch',
		)
		expect(
			coalesceRepeatCorrectnessStatus(['exact-package-match', 'semantic-roundtrip-pass']),
		).toBe('intermittent-mismatch')
	})

	test('external samples must match requested repeat count', () => {
		expect(() =>
			normalizeExternalSamples(
				{ samples: [{ durationMs: 1 }, { durationMs: 2 }] },
				3,
				'test-runner',
			),
		).toThrow('test-runner reported 2 samples but repeat requested 3')
	})

	test('external samples must provide positive finite durations', () => {
		expect(() =>
			normalizeExternalSamples({ samples: [{ durationMs: 0 }] }, 1, 'test-runner'),
		).toThrow('test-runner sample 0 must provide a positive durationMs')
		expect(() =>
			normalizeExternalSamples({ samples: [{ durationMs: Number.NaN }] }, 1, 'test-runner'),
		).toThrow('test-runner sample 0 must provide a positive durationMs')
	})

	test('external samples preserve optional memory metrics', () => {
		expect(
			normalizeExternalSamples(
				{ samples: [{ durationMs: 1, peakRssBytes: 1024, rssAfterBytes: 2048 }] },
				1,
				'test-runner',
			),
		).toEqual([{ durationMs: 1, peakRssBytes: 1024, rssAfterBytes: 2048 }])
		expect(() =>
			normalizeExternalSamples(
				{ samples: [{ durationMs: 1, peakRssBytes: Number.POSITIVE_INFINITY }] },
				1,
				'test-runner',
			),
		).toThrow('External runner sample field "peakRssBytes" must be a non-negative finite number')
	})

	test('external sample assertions are normalized for each repeat', () => {
		expect(
			normalizeExternalSampleAssertions(
				{
					assertionsBySample: [{ cellCount: 1 }, { cellCount: 2 }],
				},
				2,
				'test-runner',
			),
		).toEqual([{ cellCount: 1 }, { cellCount: 2 }])
		expect(
			normalizeExternalSampleAssertions(
				{
					samples: [
						{ durationMs: 1, assertions: { runnerVersion: '1.0.0' } },
						{ durationMs: 2, assertions: { runnerVersion: '1.0.0' } },
					],
				},
				2,
				'test-runner',
			),
		).toEqual([{ runnerVersion: '1.0.0' }, { runnerVersion: '1.0.0' }])
		expect(() =>
			normalizeExternalSampleAssertions(
				{ assertionsBySample: [{ cellCount: 1 }] },
				2,
				'test-runner',
			),
		).toThrow('test-runner reported 1 assertion samples but repeat requested 2')
		expect(() =>
			normalizeExternalSampleAssertions(
				{ assertionsBySample: [{ cellCount: {} }] },
				1,
				'test-runner',
			),
		).toThrow('External runner assertion "cellCount" must be a primitive value')
	})

	test('external assertions reject non-primitive values', () => {
		expect(() => normalizeAssertions({ assertions: { sheetCount: { value: 1 } } })).toThrow(
			'External runner assertion "sheetCount" must be a primitive value',
		)
	})

	test('runner provenance is promoted into benchmark dimensions', () => {
		expect(
			benchmarkProvenanceDimensions(
				{
					runnerVersion: '0.34.0',
					libraryVersion: '1.2.3',
					calamineVersion: '0.34.0',
					runnerEngine: 'calamine',
					cellCount: 28_056_975,
					ignoredNullVersion: null,
				},
				{
					adapterVersion: '1',
					libraryVersion: 'reported-by-runner',
					runtime: 'rust',
				},
			),
		).toEqual({
			runnerAdapterVersion: '1',
			runnerRuntime: 'rust',
			runnerVersion: '0.34.0',
			libraryVersion: '1.2.3',
			calamineVersion: '0.34.0',
			runnerEngine: 'calamine',
		})
		expect(
			benchmarkProvenanceDimensions(undefined, {
				libraryVersion: '4.4.0',
			}),
		).toEqual({ runnerManifestLibraryVersion: '4.4.0' })
	})

	test('external runner manifests reject duplicate names', () => {
		const runner = {
			name: 'runner',
			command: ['python3', 'runner.py'],
		}
		expect(() => normalizeExternalRunnerSpecs([runner, runner])).toThrow(
			'External runner "runner" is declared more than once',
		)
	})

	test('external runner manifest sets reject duplicate names across files', () => {
		const runner = {
			name: 'runner',
			command: ['python3', 'runner.py'],
		}
		expect(() => normalizeExternalRunnerManifestSet([[runner], [runner]])).toThrow(
			'External runner "runner" is declared more than once',
		)
	})

	test('external runner manifests reject malformed command and capability flags', () => {
		expect(() => normalizeExternalRunnerSpecs([{ name: 'runner', command: [] }])).toThrow(
			'External runner "runner" must provide command as a string array',
		)
		expect(() =>
			normalizeExternalRunnerSpecs([
				{
					name: 'runner',
					command: ['python3', 'runner.py'],
					capabilities: { internalTiming: 'true' },
				},
			]),
		).toThrow('External runner "runner" capability "internalTiming" must be boolean')
		expect(() =>
			normalizeExternalRunnerSpecs([
				{
					name: 'runner',
					command: ['python3', 'runner.py'],
					licenseGate: { env: '' },
				},
			]),
		).toThrow('External runner "runner" licenseGate.env must be a non-empty string')
	})

	test('external runner manifests preserve normalized metadata', () => {
		expect(
			normalizeExternalRunnerSpecs([
				{
					name: 'runner',
					command: ['python3', 'runner.py'],
					categories: ['read', 'edit-roundtrip'],
					runtime: 'python3',
					timingModel: 'external-internal-operation-timing',
					capabilities: {
						internalTiming: true,
						valueOnlyRead: true,
						metadataOnlyRead: true,
					},
				},
				{
					name: 'writer',
					command: ['writer'],
					categories: ['write'],
				},
			]),
		).toEqual([
			{
				name: 'runner',
				command: ['python3', 'runner.py'],
				categories: ['read', 'edit-roundtrip'],
				runtime: 'python3',
				timingModel: 'external-internal-operation-timing',
				capabilities: {
					internalTiming: true,
					valueOnlyRead: true,
					metadataOnlyRead: true,
				},
			},
			{
				name: 'writer',
				command: ['writer'],
				categories: ['write'],
			},
		])
	})

	test('external runner commands can override python runtime', () => {
		expect(
			resolveExternalRunnerCommand(['python3', 'runner.py', '--json'], {
				ASCEND_BENCH_PYTHON: 'fixtures/benchmarks/runners/sota_python.sh',
			}),
		).toEqual(['fixtures/benchmarks/runners/sota_python.sh', 'runner.py', '--json'])
		expect(resolveExternalRunnerCommand(['bun', 'runner.ts'])).toEqual(['bun', 'runner.ts'])
		expect(
			resolveExternalRunnerCommand(['python3', 'runner.py'], {
				ASCEND_BENCH_PYTHON: '   ',
			}),
		).toEqual(['python3', 'runner.py'])
	})

	test('rust calamine manifest follows the external runner protocol', () => {
		const parsed = JSON.parse(
			readFileSync('fixtures/benchmarks/runners/rust-calamine.manifest.json', 'utf-8'),
		) as unknown
		expect(normalizeExternalRunnerSpecs(parsed)).toEqual([
			{
				name: 'rust-calamine',
				command: [
					'cargo',
					'run',
					'--quiet',
					'--release',
					'--manifest-path',
					'fixtures/benchmarks/runners/rust-calamine/Cargo.toml',
					'--',
				],
				categories: ['read'],
				adapterVersion: '1',
				libraryVersion: 'reported-by-runner',
				runtime: 'rust',
				timingModel: 'external-internal-file-path-materialization-timing',
				validationModel: 'external-post-operation-assertions',
				memoryModel: 'peak-rss-reported',
				installHint:
					'cargo build --release --manifest-path fixtures/benchmarks/runners/rust-calamine/Cargo.toml',
				capabilities: {
					xlsmRoundtrip: false,
					internalTiming: true,
					valueOnlyRead: true,
				},
			},
		])
	})

	test('excelize manifest follows the external runner protocol', () => {
		const parsed = JSON.parse(
			readFileSync('fixtures/benchmarks/runners/excelize.manifest.json', 'utf-8'),
		) as unknown
		expect(normalizeExternalRunnerSpecs(parsed)).toEqual([
			{
				name: 'excelize',
				command: ['bash', 'fixtures/benchmarks/runners/excelize_runner.sh'],
				categories: ['read', 'write', 'edit-roundtrip'],
				adapterVersion: '1',
				libraryVersion: 'reported-by-runner',
				runtime: 'go',
				timingModel: 'external-internal-file-path-materialization-timing',
				validationModel: 'external-post-operation-assertions',
				memoryModel: 'peak-rss-reported',
				installHint: 'brew install go',
				capabilities: {
					xlsmRoundtrip: false,
					internalTiming: true,
					valueOnlyRead: true,
				},
			},
		])
	})

	test('fastexcel Java manifest follows the external runner protocol', () => {
		const parsed = JSON.parse(
			readFileSync('fixtures/benchmarks/runners/fastexcel-java.manifest.json', 'utf-8'),
		) as unknown
		expect(normalizeExternalRunnerSpecs(parsed)).toEqual([
			{
				name: 'fastexcel-java',
				command: ['bash', 'fixtures/benchmarks/runners/fastexcel_java_runner.sh'],
				categories: ['read', 'write'],
				adapterVersion: '1',
				libraryVersion: 'reported-by-runner',
				runtime: 'java',
				timingModel: 'external-internal-materialized-workbook-timing',
				validationModel: 'external-post-operation-assertions',
				memoryModel: 'jvm-heap-reported',
				installHint: 'brew install openjdk maven',
				capabilities: {
					xlsmRoundtrip: false,
					internalTiming: true,
					valueOnlyRead: true,
					finalValidation: true,
				},
			},
		])
	})

	test('ClosedXML manifest follows the external runner protocol for read and write', () => {
		const parsed = JSON.parse(
			readFileSync('fixtures/benchmarks/runners/closedxml.manifest.json', 'utf-8'),
		) as unknown
		expect(normalizeExternalRunnerSpecs(parsed)).toEqual([
			{
				name: 'closedxml',
				command: ['bash', 'fixtures/benchmarks/runners/closedxml_runner.sh'],
				categories: ['read', 'write'],
				adapterVersion: '1',
				libraryVersion: 'reported-by-runner',
				runtime: 'dotnet',
				timingModel: 'external-internal-materialized-workbook-timing',
				validationModel: 'external-post-operation-assertions',
				memoryModel: 'process-working-set-reported',
				installHint: 'brew install dotnet@8',
				capabilities: {
					xlsmRoundtrip: false,
					internalTiming: true,
					valueOnlyRead: true,
				},
			},
		])
	})

	test('NPOI manifest follows the external runner protocol', () => {
		const parsed = JSON.parse(
			readFileSync('fixtures/benchmarks/runners/npoi.manifest.json', 'utf-8'),
		) as unknown
		expect(normalizeExternalRunnerSpecs(parsed)).toEqual([
			{
				name: 'npoi',
				command: ['bash', 'fixtures/benchmarks/runners/npoi_runner.sh'],
				categories: ['write'],
				adapterVersion: '1',
				libraryVersion: 'reported-by-runner',
				runtime: 'dotnet',
				timingModel: 'external-internal-materialized-workbook-timing',
				validationModel: 'external-post-operation-assertions',
				memoryModel: 'process-working-set-reported',
				installHint:
					'brew install dotnet@8; review the NPOI OSMF license, then run with ACCEPT_NPOI_OSMF_LICENSE=1',
				licenseGate: {
					env: 'ACCEPT_NPOI_OSMF_LICENSE',
					value: '1',
					reason: 'NPOI requires explicit OSMF license acceptance.',
				},
				capabilities: {
					xlsmRoundtrip: false,
					internalTiming: true,
					valueOnlyRead: true,
				},
			},
		])
	})

	test('Polars manifest exposes each Excel engine as an explicit runner', () => {
		const parsed = JSON.parse(
			readFileSync('fixtures/benchmarks/runners/polars.manifest.json', 'utf-8'),
		) as unknown
		const specs = normalizeExternalRunnerSpecs(parsed)
		expect(specs.map((spec) => spec.name)).toEqual([
			'polars-calamine',
			'polars-xlsx2csv',
			'polars-openpyxl',
		])
		expect(specs.map((spec) => spec.command)).toEqual([
			['python3', 'fixtures/benchmarks/runners/polars_runner.py', '--engine', 'calamine'],
			['python3', 'fixtures/benchmarks/runners/polars_runner.py', '--engine', 'xlsx2csv'],
			['python3', 'fixtures/benchmarks/runners/polars_runner.py', '--engine', 'openpyxl'],
		])
		expect(specs.every((spec) => spec.capabilities?.valueOnlyRead === true)).toBe(true)
	})

	test('combined reader manifests include direct rust calamine and excelize coverage', () => {
		for (const manifestPath of [
			'fixtures/benchmarks/runners/python-readers.manifest.json',
			'fixtures/benchmarks/runners/ascend-python-readers.manifest.json',
		]) {
			const parsed = JSON.parse(readFileSync(manifestPath, 'utf-8')) as unknown
			const specs = normalizeExternalRunnerSpecs(parsed)
			const rustCalamine = specs.find((spec) => spec.name === 'rust-calamine')
			const excelize = specs.find((spec) => spec.name === 'excelize')
			const openpyxl = specs.find((spec) => spec.name === 'openpyxl')
			const fastxlsx = specs.find((spec) => spec.name === 'fastxlsx')
			const pyopenxlsx = specs.find((spec) => spec.name === 'pyopenxlsx')
			const openpyxlMetadataOnly = specs.find((spec) => spec.name === 'openpyxl-metadata-only')
			const polarsEngines = specs
				.filter((spec) => spec.name.startsWith('polars-'))
				.map((spec) => spec.name)
			expect(rustCalamine?.command).toEqual([
				'cargo',
				'run',
				'--quiet',
				'--release',
				'--manifest-path',
				'fixtures/benchmarks/runners/rust-calamine/Cargo.toml',
				'--',
			])
			expect(rustCalamine?.capabilities).toEqual({
				xlsmRoundtrip: false,
				internalTiming: true,
				valueOnlyRead: true,
			})
			expect(excelize?.command).toEqual(['bash', 'fixtures/benchmarks/runners/excelize_runner.sh'])
			expect(excelize?.categories).toEqual(['read', 'edit-roundtrip'])
			expect(excelize?.capabilities).toEqual({
				xlsmRoundtrip: false,
				internalTiming: true,
				valueOnlyRead: true,
			})
			expect(openpyxl?.categories).toEqual(['read', 'roundtrip', 'edit-roundtrip'])
			expect(openpyxlMetadataOnly?.capabilities).toEqual({
				xlsmRoundtrip: false,
				internalTiming: true,
				metadataOnlyRead: true,
			})
			expect(fastxlsx?.command).toEqual([
				'python3',
				'fixtures/benchmarks/runners/python_matrix_runner.py',
				'--library',
				'fastxlsx',
			])
			expect(pyopenxlsx?.command).toEqual([
				'python3',
				'fixtures/benchmarks/runners/python_matrix_runner.py',
				'--library',
				'pyopenxlsx',
			])
			expect(polarsEngines).toEqual(['polars-calamine', 'polars-xlsx2csv', 'polars-openpyxl'])
		}
	})

	test('corpus target selection resolves manifest filters to benchmark paths', () => {
		const entries = normalizeManifest([
			corpusEntry('pivot.xlsx', {
				pivot_tables: true,
				slicers: true,
				charts: true,
				drawings: true,
				conditional_formatting: true,
				calc_chain: true,
			}),
			corpusEntry('plain.xlsx', {}),
		])
		const selected = selectCorpusTargets(entries, { tags: ['pivot'], risks: ['high'] }, '/corpus')
		expect(selected).toEqual([
			{
				path: '/corpus/pivot.xlsx',
				corpus: {
					file: 'pivot.xlsx',
					benchmarkTier: 'extended',
					assertionClass: 'semantic-plus-package',
					riskClass: 'high',
					featureTags: [
						'calc-chain',
						'chart',
						'conditional-formatting',
						'drawing',
						'formula-fidelity',
						'pivot',
						'slicer',
						'small',
					],
					vendorable: false,
					knownUnsupported: [],
				},
			},
		])
	})

	test('default real-workbook sweeps include vendored OSS workbook coverage', () => {
		expect(QUICK_TARGETS).toContain('fixtures/xlsx/calamine/shared_formula_reversed.xlsx')
		expect(QUICK_TARGETS).toContain('fixtures/xlsx/poi/StructuredReferences.xlsx')
		expect(QUICK_TARGETS).toContain('fixtures/xlsx/poi/shared_formulas.xlsx')
		expect(QUICK_TARGETS).toContain('fixtures/xlsx/poi/WithChart.xlsx')
		expect(QUICK_TARGETS).toContain('fixtures/xlsx/libreoffice/universal-content-strict.xlsx')
		expect(QUICK_TARGETS).toContain(
			'fixtures/xlsx/libreoffice/PivotTable_CachedDefinitionAndDataInSync.xlsx',
		)
		expect(FULL_CORPUS_TARGETS).toContain('fixtures/xlsx/poi/FormulaEvalTestData_Copy.xlsx')
		expect(FULL_CORPUS_TARGETS).toContain('fixtures/xlsx/poi/NewStyleConditionalFormattings.xlsx')
		expect(FULL_CORPUS_TARGETS).toContain('fixtures/xlsx/poi/AutoFilter.xlsx')
		expect(FULL_CORPUS_TARGETS).toContain('fixtures/xlsx/poi/formula_stress_test.xlsx')
		expect(FULL_CORPUS_TARGETS).toContain('fixtures/xlsx/poi/merge_cells.xlsx')
		expect(FULL_CORPUS_TARGETS).toContain('fixtures/xlsx/poi/named_ranges_2011.xlsx')
		expect(FULL_CORPUS_TARGETS).toContain('fixtures/xlsx/libreoffice/activex_checkbox.xlsx')
		expect(FULL_CORPUS_TARGETS).toContain('fixtures/xlsx/libreoffice/textLengthDataValidity.xlsx')
		expect(FULL_CORPUS_TARGETS).toContain('fixtures/xlsx/calamine/pivots.xlsx')
		expect(FULL_CORPUS_TARGETS).toContain('fixtures/xlsx/calamine/picture.xlsx')
		expect(FULL_CORPUS_TARGETS).toContain('fixtures/xlsx/calamine/table-multiple.xlsx')
		expect(FULL_CORPUS_TARGETS).toContain(
			'fixtures/xlsx/closedxml/Misc_FormulasWithEvaluation.xlsx',
		)
		expect(FULL_CORPUS_TARGETS).toContain(
			'fixtures/xlsx/closedxml/Other_ExternalLinks_WorkbookWithExternalLink.xlsx',
		)
		expect(FULL_CORPUS_TARGETS).toContain('fixtures/xlsx/exceljs/formulas.xlsx')
		expect(FULL_CORPUS_TARGETS).toContain('fixtures/xlsx/exceljs/chart-sheet.xlsx')
	})

	test('edit-roundtrip falls back to string cells when a workbook has no numeric edit target', () => {
		const workbookPath = resolve(mkdtempSync(`${tmpdir()}/ascend-string-edit-`), 'strings.xlsx')
		writeFileSync(workbookPath, stringOnlyWorkbookBytes())
		const proc = Bun.spawnSync({
			cmd: [
				'bun',
				'run',
				'fixtures/benchmarks/competitive-real-workbook.ts',
				'--category',
				'edit-roundtrip',
				'--repeat',
				'1',
				'--warmup',
				'0',
				'--libraries',
				'ascend',
				'--json',
				workbookPath,
			],
			stdout: 'pipe',
			stderr: 'pipe',
		})
		if (proc.exitCode !== 0) {
			throw new Error(new TextDecoder().decode(proc.stderr))
		}
		const payload = JSON.parse(new TextDecoder().decode(proc.stdout)) as {
			cases: Array<{
				dimensions: { correctnessStatus: string }
				assertions: {
					editValueType?: string
					editCellValueMatches?: boolean
					semanticRoundtripMatches?: boolean
				}
			}>
			failed?: unknown[]
		}
		expect(payload.failed ?? []).toEqual([])
		expect(payload.cases[0]?.dimensions.correctnessStatus).not.toBe('error')
		expect(payload.cases[0]?.assertions.semanticRoundtripMatches).toBe(true)
		expect(payload.cases[0]?.assertions.editValueType).toBe('string')
		expect(payload.cases[0]?.assertions.editCellValueMatches).toBe(true)
	})

	test('default real-workbook sweeps span vendored OSS corpuses and feature tags', async () => {
		const physicalRoots = ['poi', 'calamine', 'closedxml', 'exceljs', 'libreoffice'] as const
		for (const root of physicalRoots) {
			expect(
				FULL_CORPUS_TARGETS.some((target) => target.startsWith(`fixtures/xlsx/${root}/`)),
			).toBe(true)
		}
		expect(FULL_CORPUS_TARGETS).toContain('fixtures/xlsx/poi/formula_stress_test.xlsx')
		expect(
			FULL_CORPUS_TARGETS.some((target) => target.startsWith('fixtures/xlsx/xlsxwriter/')),
		).toBe(true)
		expect(
			QUICK_TARGETS.some((target) =>
				/^fixtures\/xlsx\/(calamine|libreoffice|xlsxwriter)\//.test(target),
			),
		).toBe(true)

		const targetSet = new Set(FULL_CORPUS_TARGETS)
		const targetAbsSet = new Set(FULL_CORPUS_TARGETS.map((target) => resolve(target)))
		const coveredTags = new Set<string>()
		for (const root of [...physicalRoots, 'sheetjs'] as const) {
			const entries = normalizeManifest(
				await loadCorpusManifestEntries(resolve(import.meta.dir, `../xlsx/${root}/manifest.ts`)),
			)
			for (const entry of entries) {
				const covered =
					root === 'sheetjs'
						? targetAbsSet.has(resolve(import.meta.dir, `../xlsx/${root}`, entry.file))
						: targetSet.has(`fixtures/xlsx/${root}/${entry.file}`)
				if (!covered) continue
				for (const tag of entry.featureTags) coveredTags.add(tag)
			}
		}

		for (const tag of [
			'formula-fidelity',
			'pivot-table',
			'chart',
			'drawing',
			'comment',
			'conditional-formatting',
			'data-validation',
			'external-link',
			'table',
			'style',
		]) {
			expect(coveredTags.has(tag), `missing benchmark feature tag ${tag}`).toBe(true)
		}
	})

	test('TypeScript corpus manifests can promote vendored POI fixtures into benchmark selection', async () => {
		const entries = normalizeManifest(
			await loadCorpusManifestEntries(resolve(import.meta.dir, '../xlsx/poi/manifest.ts')),
		)
		if (entries.length === 0) {
			expect(
				selectCorpusTargets(
					entries,
					{ tags: ['apache-poi', 'formula-fidelity'], tiers: ['core'], vendorableOnly: true },
					resolve(import.meta.dir, '../xlsx/poi'),
				),
			).toEqual([])
			return
		}
		expect(entries.length).toBeGreaterThan(40)
		expect(validateManifestProvenance(entries)).toEqual([])
		expect(entries.some((entry) => entry.featureTags.includes('sheetjs'))).toBe(false)
		const selected = selectCorpusTargets(
			entries,
			{ tags: ['apache-poi', 'formula-fidelity'], tiers: ['core'], vendorableOnly: true },
			resolve(import.meta.dir, '../xlsx/poi'),
		)
		expect(selected.some((entry) => entry.path.endsWith('StructuredReferences.xlsx'))).toBe(true)
	})

	test('TypeScript corpus manifests promote SheetJS fixtures separately from POI', async () => {
		const entries = normalizeManifest(
			await loadCorpusManifestEntries(resolve(import.meta.dir, '../xlsx/sheetjs/manifest.ts')),
		)
		if (entries.length === 0) {
			expect(
				selectCorpusTargets(
					entries,
					{ tags: ['sheetjs', 'formula-fidelity'], tiers: ['core'], vendorableOnly: true },
					resolve(import.meta.dir, '../xlsx/sheetjs'),
				),
			).toEqual([])
			return
		}
		expect(entries.map((entry) => entry.file).sort()).toEqual([
			'../poi/AutoFilter.xlsx',
			'../poi/formula_stress_test.xlsx',
			'../poi/merge_cells.xlsx',
			'../poi/named_ranges_2011.xlsx',
		])
		expect(validateManifestProvenance(entries)).toEqual([])
		expect(entries.every((entry) => entry.featureTags.includes('sheetjs'))).toBe(true)
		expect(entries.every((entry) => !entry.featureTags.includes('apache-poi'))).toBe(true)
		for (const entry of entries) {
			expect(FULL_CORPUS_TARGETS).toContain(`fixtures/xlsx/poi/${entry.file.split('/').at(-1)}`)
		}
		const selected = selectCorpusTargets(
			entries,
			{ tags: ['sheetjs', 'formula-fidelity'], tiers: ['core'], vendorableOnly: true },
			resolve(import.meta.dir, '../xlsx/sheetjs'),
		)
		expect(selected.some((entry) => entry.path.endsWith('poi/formula_stress_test.xlsx'))).toBe(true)
	})

	test('TypeScript corpus manifests can promote vendored LibreOffice fixtures into benchmark selection', async () => {
		const entries = normalizeManifest(
			await loadCorpusManifestEntries(resolve(import.meta.dir, '../xlsx/libreoffice/manifest.ts')),
		)
		expect(entries.length).toBe(22)
		expect(validateManifestProvenance(entries)).toEqual([])
		const selected = selectCorpusTargets(
			entries,
			{ tags: ['libreoffice', 'pivot-table'], tiers: ['core'], vendorableOnly: true },
			resolve(import.meta.dir, '../xlsx/libreoffice'),
		)
		expect(
			selected.some((entry) =>
				entry.path.endsWith('PivotTable_CachedDefinitionAndDataInSync.xlsx'),
			),
		).toBe(true)
	})

	test('feature summary extracts package and worksheet feature inventory', () => {
		const summary = extractWorkbookFeatureSummary(featureWorkbookBytes())
		expect(summary.tablePartCount).toBe(1)
		expect(summary.chartPartCount).toBe(1)
		expect(summary.drawingPartCount).toBe(1)
		expect(summary.vmlDrawingPartCount).toBe(1)
		expect(summary.pivotTablePartCount).toBe(1)
		expect(summary.pivotCachePartCount).toBe(1)
		expect(summary.commentPartCount).toBe(1)
		expect(summary.mediaPartCount).toBe(1)
		expect(summary.externalLinkPartCount).toBe(1)
		expect(summary.connectionPartCount).toBe(1)
		expect(summary.customXmlPartCount).toBe(1)
		expect(summary.worksheetHyperlinkCount).toBe(1)
		expect(summary.worksheetDataValidationCount).toBe(1)
		expect(summary.worksheetConditionalFormattingCount).toBe(1)
		expect(summary.definedNameCount).toBe(1)
		expect(summary.featurePartNamesHash).not.toBe(emptyHash())
		expect(summary.featureInventoryHash).not.toBe(emptyHash())
	})

	test('feature summary hash changes when worksheet features are dropped', () => {
		const withFeature = extractWorkbookFeatureSummary(featureWorkbookBytes())
		const withoutConditionalFormatting = extractWorkbookFeatureSummary(
			featureWorkbookBytes({ conditionalFormatting: false }),
		)
		expect(withoutConditionalFormatting.worksheetConditionalFormattingCount).toBe(0)
		expect(withoutConditionalFormatting.featurePartNamesHash).toBe(withFeature.featurePartNamesHash)
		expect(withoutConditionalFormatting.featureInventoryHash).not.toBe(
			withFeature.featureInventoryHash,
		)
	})

	test('feature summary includes nested comment part paths in feature hashes', () => {
		const legacyComments = extractWorkbookFeatureSummary(
			featureWorkbookBytes({ commentPath: 'xl/comments1.xml' }),
		)
		const nestedComments = extractWorkbookFeatureSummary(
			featureWorkbookBytes({ commentPath: 'xl/comments/comment1.xml' }),
		)
		expect(nestedComments.commentPartCount).toBe(1)
		expect(nestedComments.featurePartNamesHash).not.toBe(emptyHash())
		expect(nestedComments.featurePartNamesHash).not.toBe(legacyComments.featurePartNamesHash)
		expect(nestedComments.featureInventoryHash).not.toBe(legacyComments.featureInventoryHash)
	})

	test('feature summary hash changes when embedded worksheet features are dropped', () => {
		const withFeature = extractWorkbookFeatureSummary(
			featureWorkbookBytes({
				autoFilter: true,
				calcChain: true,
				mergeCells: true,
				richSharedString: true,
				sheetProtection: true,
				sheetView: true,
			}),
		)
		const withoutMerge = extractWorkbookFeatureSummary(
			featureWorkbookBytes({
				autoFilter: true,
				calcChain: true,
				mergeCells: false,
				richSharedString: true,
				sheetProtection: true,
				sheetView: true,
			}),
		)
		expect(withoutMerge.featurePartNamesHash).toBe(withFeature.featurePartNamesHash)
		expect(withoutMerge.featureInventoryHash).not.toBe(withFeature.featureInventoryHash)
	})
})

function shape(overrides: Partial<WorkbookShapeSummary> = {}): WorkbookShapeSummary {
	return {
		sheetNames: ['Sheet1'],
		sheetCount: 1,
		cellCount: 2,
		physicalCellCount: 2,
		formulaCount: 1,
		usedRanges: ['Sheet1!A1:B1'],
		physicalUsedRanges: ['Sheet1!A1:B1'],
		sheetNamesHash: 'sheet-names',
		usedRangesHash: 'used-ranges',
		physicalUsedRangesHash: 'physical-used-ranges',
		semanticCellRefsHash: 'semantic-refs',
		semanticCellValuesHash: 'semantic-values',
		formulaTextHash: 'formulas',
		...overrides,
	}
}

function packageFingerprint(
	overrides: Partial<WorkbookPackageFingerprint> = {},
): WorkbookPackageFingerprint {
	return {
		partCount: 6,
		partNamesHash: 'package-parts',
		contentTypeCount: 4,
		contentTypesHash: 'content-types',
		relationshipCount: 3,
		relationshipGraphHash: 'relationship-graph',
		preservedPartCount: 1,
		preservedPartNamesHash: 'preserved-parts',
		preservedPartContentHash: 'preserved-content',
		...overrides,
	}
}

function featureSummary(overrides: Partial<WorkbookFeatureSummary> = {}): WorkbookFeatureSummary {
	return {
		tablePartCount: 1,
		chartPartCount: 1,
		chartExPartCount: 0,
		drawingPartCount: 1,
		vmlDrawingPartCount: 0,
		pivotTablePartCount: 1,
		pivotCachePartCount: 2,
		slicerPartCount: 0,
		commentPartCount: 1,
		threadedCommentPartCount: 0,
		mediaPartCount: 1,
		externalLinkPartCount: 0,
		connectionPartCount: 0,
		customXmlPartCount: 0,
		worksheetHyperlinkCount: 2,
		worksheetDataValidationCount: 1,
		worksheetConditionalFormattingCount: 1,
		definedNameCount: 1,
		featurePartNamesHash: 'feature-parts',
		featureInventoryHash: 'feature-inventory',
		...overrides,
	}
}

function passingReadAssertions(
	expected: WorkbookShapeSummary,
): Record<string, string | number | boolean | null> {
	return {
		sheetCount: expected.sheetCount,
		sheetNamesHash: expected.sheetNamesHash,
		cellCount: expected.cellCount,
		formulaCount: expected.formulaCount,
		firstUsedRange: expected.usedRanges[0] ?? null,
		usedRangesHash: expected.usedRangesHash,
		physicalUsedRangesHash: expected.physicalUsedRangesHash,
		semanticCellRefsHash: expected.semanticCellRefsHash,
		semanticCellValuesHash: expected.semanticCellValuesHash,
		formulaTextHash: expected.formulaTextHash,
	}
}

function passingRoundtripAssertions(
	expected: WorkbookShapeSummary,
): Record<string, string | number | boolean | null> {
	const packageInfo = expected.packageFingerprint
	const features = expected.featureSummary
	return {
		byteIdentical: false,
		roundtripSheetCount: expected.sheetCount,
		roundtripSheetNamesHash: expected.sheetNamesHash,
		roundtripCellCount: expected.cellCount,
		roundtripFormulaCount: expected.formulaCount,
		roundtripFirstUsedRange: expected.usedRanges[0] ?? null,
		roundtripUsedRangesHash: expected.usedRangesHash,
		roundtripSemanticCellRefsHash: expected.semanticCellRefsHash,
		roundtripSemanticCellValuesHash: expected.semanticCellValuesHash,
		roundtripFormulaTextHash: expected.formulaTextHash,
		roundtripPackagePartCount: packageInfo?.partCount ?? null,
		roundtripPackagePartNamesHash: packageInfo?.partNamesHash ?? null,
		roundtripPackageContentTypeCount: packageInfo?.contentTypeCount ?? null,
		roundtripPackageContentTypesHash: packageInfo?.contentTypesHash ?? null,
		roundtripPackageRelationshipCount: packageInfo?.relationshipCount ?? null,
		roundtripPackageRelationshipGraphHash: packageInfo?.relationshipGraphHash ?? null,
		roundtripPreservedPartCount: packageInfo?.preservedPartCount ?? null,
		roundtripPreservedPartNamesHash: packageInfo?.preservedPartNamesHash ?? null,
		roundtripPreservedPartContentHash: packageInfo?.preservedPartContentHash ?? null,
		...(features
			? {
					roundtripTablePartCount: features.tablePartCount,
					roundtripChartPartCount: features.chartPartCount,
					roundtripChartExPartCount: features.chartExPartCount,
					roundtripDrawingPartCount: features.drawingPartCount,
					roundtripVmlDrawingPartCount: features.vmlDrawingPartCount,
					roundtripPivotTablePartCount: features.pivotTablePartCount,
					roundtripPivotCachePartCount: features.pivotCachePartCount,
					roundtripSlicerPartCount: features.slicerPartCount,
					roundtripCommentPartCount: features.commentPartCount,
					roundtripThreadedCommentPartCount: features.threadedCommentPartCount,
					roundtripMediaPartCount: features.mediaPartCount,
					roundtripExternalLinkPartCount: features.externalLinkPartCount,
					roundtripConnectionPartCount: features.connectionPartCount,
					roundtripCustomXmlPartCount: features.customXmlPartCount,
					roundtripWorksheetHyperlinkCount: features.worksheetHyperlinkCount,
					roundtripWorksheetDataValidationCount: features.worksheetDataValidationCount,
					roundtripWorksheetConditionalFormattingCount:
						features.worksheetConditionalFormattingCount,
					roundtripDefinedNameCount: features.definedNameCount,
					roundtripFeaturePartNamesHash: features.featurePartNamesHash,
					roundtripFeatureInventoryHash: features.featureInventoryHash,
				}
			: {}),
	}
}

function corpusEntry(
	file: string,
	features: Partial<Record<string, boolean>>,
): CorpusManifestEntry {
	return {
		file,
		size_bytes: 100_000,
		features: {
			macros: false,
			charts: false,
			pivot_tables: false,
			tables: false,
			drawings: false,
			comments: false,
			threaded_comments: false,
			conditional_formatting: false,
			data_validations: false,
			merged_cells: false,
			hyperlinks: false,
			defined_names: false,
			external_links: false,
			connections: false,
			slicers: false,
			images_or_media: false,
			custom_xml: false,
			calc_chain: false,
			...features,
		},
		counts: {
			worksheets: 1,
			charts: 0,
			tables: 0,
			drawings: 0,
			pivot_tables: 0,
			pivot_caches: 0,
			comments: 0,
		},
	}
}

function emptyHash(): string {
	return 'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855'
}

function stringOnlyWorkbookBytes(): Uint8Array {
	return zipSync({
		'[Content_Types].xml': strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
		'_rels/.rels': strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
		'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
		'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
		'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>alpha</t></is></c></row></sheetData>
</worksheet>`),
	})
}

function featureWorkbookBytes(
	options: {
		readonly autoFilter?: boolean
		readonly calcChain?: boolean
		readonly commentPath?: 'xl/comments1.xml' | 'xl/comments/comment1.xml'
		readonly conditionalFormatting?: boolean
		readonly mergeCells?: boolean
		readonly richSharedString?: boolean
		readonly sheetProtection?: boolean
		readonly sheetView?: boolean
	} = {},
): Uint8Array {
	const includeConditionalFormatting = options.conditionalFormatting !== false
	const commentPath = options.commentPath ?? 'xl/comments1.xml'
	return zipSync({
		'[Content_Types].xml': strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  ${
		options.richSharedString
			? '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
			: ''
	}
  ${
		options.calcChain
			? '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>'
			: ''
	}
</Types>`),
		'_rels/.rels': strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
		'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
		'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <definedNames><definedName name="SalesRange">Sheet1!$A$1:$B$2</definedName></definedNames>
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
		'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  ${
		options.sheetView
			? '<sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane state="frozen" ySplit="1" topLeftCell="A2"/></sheetView></sheetViews>'
			: ''
	}
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
  ${options.mergeCells ? '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>' : ''}
  ${
		options.autoFilter
			? '<autoFilter ref="A1:B3"><sortState ref="A2:B3"><sortCondition ref="B2:B3"/></sortState></autoFilter>'
			: ''
	}
  <hyperlinks><hyperlink ref="A1" r:id="rIdHyperlink"/></hyperlinks>
  <dataValidations count="1"><dataValidation type="whole" sqref="A1"/></dataValidations>
  ${
		includeConditionalFormatting
			? '<conditionalFormatting sqref="A1"><cfRule type="cellIs" priority="1"/></conditionalFormatting>'
			: ''
	}
  ${
		options.sheetProtection
			? '<sheetProtection sheet="1" objects="1" scenarios="1" autoFilter="0"/>'
			: ''
	}
</worksheet>`),
		'xl/tables/table1.xml': strToU8('<table/>'),
		'xl/charts/chart1.xml': strToU8('<c:chartSpace/>'),
		'xl/drawings/drawing1.xml': strToU8('<xdr:wsDr/>'),
		'xl/drawings/vmlDrawing1.vml': strToU8('<xml/>'),
		'xl/pivotTables/pivotTable1.xml': strToU8('<pivotTableDefinition/>'),
		'xl/pivotCache/pivotCacheDefinition1.xml': strToU8('<pivotCacheDefinition/>'),
		[commentPath]: strToU8('<comments/>'),
		'xl/media/image1.png': new Uint8Array([137, 80, 78, 71]),
		'xl/externalLinks/externalLink1.xml': strToU8('<externalLink/>'),
		'xl/connections.xml': strToU8('<connections/>'),
		'customXml/item1.xml': strToU8('<root/>'),
		...(options.richSharedString
			? {
					'xl/sharedStrings.xml': strToU8(
						'<sst><si><r><t>Rich</t></r><r><t> Text</t></r></si></sst>',
					),
				}
			: {}),
		...(options.calcChain
			? {
					'xl/calcChain.xml': strToU8('<calcChain><c r="B1" i="1"/></calcChain>'),
				}
			: {}),
	})
}
