#!/bin/sh
set -eu

SCRIPT_DIR="$(CDPATH= cd -- "$(dirname "$0")" && pwd)"
REPO_ROOT="$(CDPATH= cd -- "$SCRIPT_DIR/../../.." && pwd)"

sh "$SCRIPT_DIR/download.sh"
cd "$REPO_ROOT"
bun test fixtures/xlsx/xlsx-fixtures.test.ts
