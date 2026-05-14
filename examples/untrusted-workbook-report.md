# Untrusted Workbook Report

Use this before reading workbook text into an agent prompt, especially for files from email, downloads, customers, or another agent.

## CLI

```bash
ascend inspect vendor-model.xlsx --agent --json
```

Expected JSON fields:

```json
{
  "ok": true,
  "data": {
    "trust": "untrusted",
    "posture": "safe-parser-preserver",
    "includedInAgentContext": {
      "visibleSheets": true,
      "hiddenSheets": false,
      "comments": false,
      "definedNames": false,
      "externalContent": false,
      "activeContent": false
    },
    "executionPolicy": {
      "macros": "preserve-only",
      "activeX": "preserve-only",
      "oleObjects": "preserve-only",
      "dde": "do-not-execute",
      "externalLinks": "do-not-refresh",
      "dataConnections": "do-not-refresh",
      "formulas": "pure-evaluation-only"
    },
    "findings": [
      {
        "code": "workbook.vbaProject",
        "severity": "warning",
        "category": "active-content",
        "location": { "partPath": "xl/vbaProject.bin" },
        "nextAction": "Do not execute macros; preserve the VBA project only if the user approves edits to this workbook."
      }
    ],
    "nextActions": [
      "Read only the visible ranges needed for the task.",
      "Use plan before commit and verify the output workbook."
    ]
  }
}
```

The report is not a numeric risk score. It is a boundary map: what is safe to include in default agent context, what remains preserve-only, where suspicious or hidden content lives, and what to do next.

## HTTP

```bash
curl -s http://localhost:3000/trust-report \
  -H 'content-type: application/json' \
  -d '{"file":"vendor-model.xlsx","maxFindings":50}'
```

## MCP

```json
{
  "tool": "ascend.trust_report",
  "arguments": {
    "file": "vendor-model.xlsx",
    "maxFindings": 50
  }
}
```

After the trust report, continue the normal safe edit workflow: inspect/read the needed visible data, build operations, plan, commit to a new output path, then check/lint/diff/trace or repair.
