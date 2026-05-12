import type { WritePolicyPackagePart, WritePolicyReport } from '@ascend/sdk'
import { bullet } from '../output/pretty.ts'

export function printWritePolicySummary(writePolicy: WritePolicyReport): void {
	console.log(bullet('Write policy diagnostics', writePolicy.diagnostics.length))
	for (const diagnostic of writePolicy.diagnostics.filter((entry) => entry.severity !== 'info')) {
		console.log(
			bullet(`Write policy ${diagnostic.code}`, `${diagnostic.severity}: ${diagnostic.message}`),
		)
		if (diagnostic.packageParts && diagnostic.packageParts.length > 0) {
			console.log(
				bullet(
					`${diagnostic.code} package parts`,
					formatPackagePartSummary(diagnostic.packageParts),
				),
			)
		}
	}
}

function formatPackagePartSummary(parts: readonly WritePolicyPackagePart[]): string {
	const visible = parts
		.slice(0, 6)
		.map(
			(part) =>
				`${part.partPath} (${part.featureFamily}, ${part.ownerScope}, ${part.preservationPolicy})`,
		)
	const remaining = parts.length - visible.length
	if (remaining > 0) visible.push(`+${remaining} more`)
	return visible.join('; ')
}
