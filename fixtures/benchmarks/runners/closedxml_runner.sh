#!/usr/bin/env bash
set -euo pipefail

tmp_root="${TMPDIR:-/tmp}"
if [[ -z "${DOTNET_ROOT:-}" && -d /opt/homebrew/opt/dotnet@8/libexec ]]; then
	export DOTNET_ROOT=/opt/homebrew/opt/dotnet@8/libexec
fi
if [[ -d /opt/homebrew/opt/dotnet@8/bin ]]; then
	export PATH="/opt/homebrew/opt/dotnet@8/bin:$PATH"
fi
export DOTNET_CLI_HOME="${DOTNET_CLI_HOME:-$tmp_root/ascend-dotnet-home}"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE="${DOTNET_SKIP_FIRST_TIME_EXPERIENCE:-1}"
export DOTNET_NOLOGO="${DOTNET_NOLOGO:-1}"
export NUGET_PACKAGES="${NUGET_PACKAGES:-$tmp_root/ascend-nuget-packages}"

exec dotnet run --project fixtures/benchmarks/runners/closedxml/ClosedXmlRunner.csproj \
	--configuration Release -- "$@"
