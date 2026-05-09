import type { CheckIssue as VerifyCheckIssue } from '@ascend/verify'
import type { CheckIssue } from './types.ts'

export function sdkCheckIssueFromVerify(issue: VerifyCheckIssue): CheckIssue {
	return {
		rule: issue.rule,
		severity: issue.severity,
		message: issue.message,
		...(issue.refs?.[0] ? { ref: issue.refs[0] } : {}),
		...(issue.refs ? { refs: issue.refs } : {}),
		...(issue.suggestedFix ? { suggestedFix: issue.suggestedFix } : {}),
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
