import type { CheckIssue as VerifyCheckIssue } from '@ascend/verify'
import type { CheckIssue, LintWarning } from './types.ts'

export function sdkCheckIssueFromVerify(issue: VerifyCheckIssue): CheckIssue {
	return {
		rule: issue.rule,
		severity: issue.severity,
		message: issue.message,
		...(issue.refs?.[0] ? { ref: issue.refs[0] } : {}),
		...(issue.refs ? { refs: issue.refs } : {}),
		...(issue.suggestedFix ? { suggestedFix: issue.suggestedFix } : {}),
		...(issue.details ? { details: issue.details } : {}),
	}
}

export function partialDependencyCheckIssue(message: string): CheckIssue {
	return {
		rule: 'partial-dependency-analysis',
		severity: 'warning',
		message,
		suggestedFix: 'Open the workbook with all referenced sheets loaded before running check.',
	}
}

export function partialDependencyLintWarning(message: string): LintWarning {
	return {
		rule: 'partial-dependency-analysis',
		severity: 'warning',
		message,
	}
}
