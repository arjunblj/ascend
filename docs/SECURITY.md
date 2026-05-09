# Security

## Formula evaluation and codegen

Ascend may compile hot formula paths using JavaScript `new Function(...)` (see `packages/engine/src/codegen.ts`). Generated code runs in the same process as the host application.

**Treat workbook files and formula text as trusted input** unless you isolate evaluation (separate process, sandboxed worker, or disallowing file uploads of untrusted `.xlsx`).

## HTTP API (`apps/api`)

The reference server accepts **local file paths** in JSON bodies. Do not expose it to the public internet without authentication and path allowlisting.

## Dependencies

Keep third-party packages updated. Run `bun audit` periodically.
