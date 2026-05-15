# Table Creation Validation Boundary

## Question

Should table creation reject invalid public metadata before mutating workbook table state, and should journal exactness classify the same cases?

## Hypothesis

Yes. Table creation is a high-impact metadata operation: it defines a range, name, headers, and table column metadata. Invalid public values should fail before table insertion and should be visible in journal exactness.

## External sources checked

- Microsoft Excel table overview: https://support.microsoft.com/en-us/office/overview-of-excel-tables-7ab0bb7d-3a9e-4b56-a3c9-6c94334e492c
- Microsoft Excel create and format tables: https://support.microsoft.com/en-us/office/create-and-format-tables-e81aa349-b006-4f8a-9806-5af9df0ac664
- OOXML table definition reference: https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_table_topic_ID0EUBU5.html

## Why this matters to Ascend

Ascend's auditable mutation claim depends on table operations being reject-first when public inputs are invalid. Creating a malformed table would affect package writes, formula references, range metadata, and later table mutations, so the boundary belongs in the correctness loop.

## Probe/implementation

An in-flight fix landed as `993ea086 fix(engine): reject invalid table creation metadata`.

It added:

- engine validation for invalid table creation metadata.
- regression coverage proving invalid table creation inputs reject.
- SDK journal exactness classification for the same invalid values as unsupported metadata.

## Results

Validation run while this diff was present:

- `bun run test:changed` passed with 5200 pass, 1 skip, 0 fail.

The commit changed:

- `packages/engine/src/operations/table-ops.ts`
- `packages/engine/src/operations.test.ts`
- `packages/sdk/src/journal.ts`
- `packages/sdk/src/journal-exactness.test.ts`

## Confidence

Medium-high. The fix is already committed and full changed tests passed, but this log did not rerun a separate isolated focused command after the commit because the worktree still contains unrelated RC/IO reader edits.

## Fold-in decision

Promote to correctness loop as auditable mutation hygiene. Do not promote new table product claims.

## Next question

The remaining dirty files are release RC gate and IO reader full-scalar parsing work. Validate them as separate release/performance candidates only if they stay coherent and isolated.
