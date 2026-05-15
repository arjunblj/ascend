# Journal Issue Registry Proof

## Question

Do generated mutation journal issues always use the exported v1 issue code, surface, and reason registries that agents and release reports depend on?

## Hypothesis

A small SDK regression assertion can prove every generated lossy journal issue is classified through the public registry arrays, not just through ad hoc string presence checks.

## External sources checked

- Language Server Protocol diagnostics use stable machine-readable diagnostic fields, including codes, for tool interoperability: https://microsoft.github.io/language-server-protocol/specifications/lsp/3.17/specification/
- OpenTelemetry semantic conventions define shared names and attributes so telemetry can be interpreted consistently across tools: https://opentelemetry.io/docs/specs/otel/semantic-conventions/
- SLSA provenance uses structured predicate fields for downstream verification, reinforcing that local release proof must keep structured vocabularies stable even when it is not signed provenance: https://slsa.dev/spec/v1.0-rc1/provenance

## Why this matters to Ascend

The auditable package-part mutation claim relies on agent-readable journal issues. If issue codes, surfaces, or reasons drift outside the exported vocabulary, downstream CLI/API/MCP summaries and release proof reports can become impossible to compare or validate.

## Probe/implementation

Finished the in-flight SDK test change in `packages/sdk/src/journal-exactness.test.ts`:

- imported `MUTATION_JOURNAL_ISSUE_CODES`;
- imported `MUTATION_JOURNAL_SURFACES`;
- imported `MUTATION_JOURNAL_REASON_CODES`;
- asserted generated entry-level issues and journal-level issues use only exported codes, surfaces, and reasons before classification.

Commands run:

```bash
bun test packages/sdk/src/journal-exactness.test.ts
bunx biome check packages/sdk/src/journal-exactness.test.ts
bunx tsc --build
```

## Results

- `journal-exactness.test.ts` passed: 33 tests, 2136 assertions.
- Biome passed for the touched SDK test file.
- `bunx tsc --build` passed.

## Confidence

High. The assertion covers generated lossy journals and checks the registry arrays directly before calling the classifier.

## Fold-in decision

Promote to correctness loop. This is test-only hardening for the auditable mutation claim and does not add a new user-facing surface.

## Next question

Should the package-action proof harness expose a compact count of journal issue codes by surface and reason, or is that too close to another release report surface before the publication policy is approved?
