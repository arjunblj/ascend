#!/usr/bin/env bash
set -euo pipefail

tmp_root="${TMPDIR:-/tmp}"
export GOMODCACHE="${GOMODCACHE:-$tmp_root/ascend-go-mod-cache}"
export GOCACHE="${GOCACHE:-$tmp_root/ascend-go-build-cache}"

ROOT="$(pwd)"
ARGS=()
while (($#)); do
	if [[ "$1" == "--file" && $# -ge 2 ]]; then
		ARGS+=("$1")
		shift
		case "$1" in
			/*) ARGS+=("$1") ;;
			*) ARGS+=("$ROOT/$1") ;;
		esac
	else
		ARGS+=("$1")
	fi
	shift
done

exec go -C fixtures/benchmarks/runners/excelize run . "${ARGS[@]}"
