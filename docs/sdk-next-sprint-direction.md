# SDK next-sprint direction (concise)

The next sprint should keep the write UX centered on reviewable intent, not direct mutation.

## Core flow

Prefer this path as the primary story:

`open -> select -> propose -> review -> commit -> save`

- **open**: load in the lightest mode that supports the task.
- **select**: work from explicit ranges/refs instead of implicit global mutation.
- **propose**: produce operation intents first.
- **review**: surface user-visible diffs and capability limits before committing.
- **commit**: apply accepted changes.
- **save**: persist bytes after successful review/commit.

## Keep review-first as default

- `preview()` / `review()` remain the center of write UX.
- Direct apply paths stay available, but should not be the recommended first step.
- Any new high-level helpers should compose to preview-first behavior.

## Preserve partial-view honesty

- Keep explicit partial-view capability errors from `packages/sdk/src/read-view.ts`.
- Do not silently pretend a partial load is full-fidelity.
- Continue to prefer explicit user-facing diagnostics over hidden fallback behavior.

## Keep advanced escape hatch, not the default

- `Operation[]` remains the advanced low-level escape hatch.
- High-level SDK affordances should compile down to `Operation[]` internally.
- Documentation and examples should lead with review-first workflows, then show `Operation[]` for expert cases.
