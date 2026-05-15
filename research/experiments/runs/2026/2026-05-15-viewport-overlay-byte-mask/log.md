# Viewport Overlay Byte Mask

## Question

Can interactive viewport overlay flags use a dense byte mask instead of `Set<number>` offsets without changing viewport semantics?

## Hypothesis

Yes. Viewport cells already have a stable flat index. For a bounded viewport, a `Uint8Array` mask should preserve `has(offset)` behavior while avoiding per-cell boxed number entries.

## External sources checked

- MDN `Uint8Array`: https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Uint8Array
- MDN typed arrays guide: https://developer.mozilla.org/en-US/docs/Web/JavaScript/Guide/Typed_arrays
- MDN `Set.prototype.has`: https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Set/has

## Why this matters to Ascend

Retained viewport patch history and agent UI reads are proof-backed claims, but they should remain cheap for dense operational views. Overlay flags for merges, validations, conditional formats, and tables are read for every visible cell, so they are a good place to remove avoidable allocation without changing surface semantics.

## Probe/implementation

- Replaced `rangeMaskOffsets` usage in interactive viewport range context with a local `viewportRangeMask`.
- `viewportRangeMask` intersects each metadata range with the viewport and writes `1` into a `Uint8Array` by flat viewport offset.
- Preserved a `.has(offset)` interface so downstream cell materialization remains unchanged.
- Kept an empty singleton mask for viewports with no relevant metadata ranges.

## Results

Validation:

```bash
bun test packages/sdk/src/interactive-contract.test.ts -t "viewport"
bun test fixtures/benchmarks/agent-workflow.test.ts
bun test fixtures/benchmarks/agent-phase-profile.test.ts
bunx biome check packages/sdk/src/session.ts
bunx tsc --build
bun run test:changed
```

`bun run test:changed` passed with `5166 pass`, `1 skip`, `0 fail` across `186` files.

## Confidence

High for semantic parity on current viewport contract tests. Medium for performance until a targeted viewport benchmark measures allocation and latency on large metadata-dense views.

## Fold-in decision

Promote to performance loop as a small implementation. Do not promote a product claim beyond existing retained viewport patch/overlay behavior; this is internal allocation hygiene.

## Next question

Can the viewport proof harness report metadata-overlay allocation or mask-build time so future UI/agent performance claims have a measurable gate?
