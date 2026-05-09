#!/usr/bin/env bash
set -euo pipefail

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/../../.." && pwd)"
venv="${ASCEND_SOTA_PYTHON_VENV:-/private/tmp/ascend-sota-python}"
python_version="${ASCEND_SOTA_PYTHON_VERSION:-3.12}"
requirements="$repo_root/fixtures/benchmarks/runners/requirements-sota.txt"
stamp="$venv/.requirements-sota.sha256"

if ! command -v uv >/dev/null 2>&1; then
	exec python3 "$@"
fi

requirements_hash="$(shasum -a 256 "$requirements" | awk '{print $1}')"
if [[ ! -x "$venv/bin/python" || ! -f "$stamp" || "$(cat "$stamp")" != "$requirements_hash" ]]; then
	uv venv --python "$python_version" "$venv"
	uv pip install --python "$venv/bin/python" -r "$requirements"
	printf '%s\n' "$requirements_hash" > "$stamp"
fi

exec "$venv/bin/python" "$@"
