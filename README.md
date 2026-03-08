# Ascend

Agent-native spreadsheet platform. Read, edit, verify, and export `.xlsx` workbooks from TypeScript.

## Quickstart

```bash
bun add ascend
```

```typescript
import { Ascend } from 'ascend'

const wb = await Ascend.open('model.xlsx')
const sheet = wb.sheet('Revenue')
const data = sheet.range('A1:D10')

wb.check()  // structural integrity
wb.lint()   // formula quality
wb.trace('Revenue!E2')  // dependency tree

await wb.save('output.xlsx')
```

## Development

```bash
bun install
bun test --recursive
bunx biome check
bunx tsc --build
```

## License

Apache-2.0
