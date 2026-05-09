# Versioning and publishing

- **Semver:** Use `MAJOR.MINOR.PATCH` for published `@ascend/*` packages.
- **Workspace:** Packages are `private: true` during monorepo development; flip `private` and set a real version before npm publish.
- **Build:** Run `bunx tsc --build` before release so declaration outputs under each package `dist/` are fresh (SDK uses `emitDeclarationOnly` in its default `tsconfig.json`).
- **Scope:** Prefer scoped names (`@ascend/sdk`, `@ascend/core`, …) for clarity on npm.

## Monorepo strategies

1. **Publish all packages** with aligned versions (changesets or manual bump).
2. **Publish SDK only** and document that consumers must use Bun/TypeScript path aliases to sibling packages until the full graph is published.

Choose based on how much of the stack external apps should depend on directly.
