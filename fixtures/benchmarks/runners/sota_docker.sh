#!/usr/bin/env bash
set -euo pipefail

image="${ASCEND_SOTA_IMAGE:-ascend-sota-bench:local}"
repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/../../.." && pwd)"
command="${1:-run}"

build_image() {
	docker build -f "$repo_root/fixtures/benchmarks/runners/Dockerfile.sota" -t "$image" "$repo_root"
}

run_image() {
	if ! docker image inspect "$image" >/dev/null 2>&1; then
		build_image
	fi

	local -a tty_args=()
	if [[ -t 0 && -t 1 ]]; then
		tty_args=(-it)
	fi
	local -a resource_args=()
	if [[ -n "${ASCEND_SOTA_CPUSET:-}" ]]; then
		resource_args+=(--cpuset-cpus "$ASCEND_SOTA_CPUSET")
	fi
	if [[ -n "${ASCEND_SOTA_MEMORY:-}" ]]; then
		resource_args+=(--memory "$ASCEND_SOTA_MEMORY")
	fi

	if [[ $# -eq 0 ]]; then
		set -- bash
	fi

	local -a docker_args=(
		--rm
		--workdir /workspace
		--mount "type=bind,src=$repo_root,target=/workspace"
		--mount "type=volume,src=ascend-sota-node-modules,target=/workspace/node_modules"
		--mount "type=volume,src=ascend-sota-bun-cache,target=/root/.bun/install/cache"
		--mount "type=volume,src=ascend-sota-cargo,target=/tmp/ascend-cargo-home"
		--mount "type=volume,src=ascend-sota-cargo-target,target=/tmp/ascend-cargo-target"
		--mount "type=volume,src=ascend-sota-go,target=/go"
		--mount "type=volume,src=ascend-sota-maven,target=/tmp/ascend-m2-repository"
		--mount "type=volume,src=ascend-sota-nuget,target=/tmp/ascend-nuget-packages"
		--mount "type=volume,src=ascend-sota-dotnet-home,target=/tmp/ascend-dotnet-home"
		--env MAVEN_REPO_LOCAL=/tmp/ascend-m2-repository
		--env NUGET_PACKAGES=/tmp/ascend-nuget-packages
		--env DOTNET_CLI_HOME=/tmp/ascend-dotnet-home
		--env DOTNET_CLI_TELEMETRY_OPTOUT=1
		--env DOTNET_NOLOGO=1
		--env CARGO_HOME=/tmp/ascend-cargo-home
		--env CARGO_TARGET_DIR=/tmp/ascend-cargo-target
		--env GOMODCACHE=/go/pkg/mod
		--env GOCACHE=/go/build-cache
		--env PATH=/usr/local/go/bin:/usr/local/cargo/bin:/usr/local/sbin:/usr/local/bin:/usr/sbin:/usr/bin:/sbin:/bin
	)
	if [[ ${#tty_args[@]} -gt 0 ]]; then
		docker_args+=("${tty_args[@]}")
	fi
	if [[ ${#resource_args[@]} -gt 0 ]]; then
		docker_args+=("${resource_args[@]}")
	fi

	docker run "${docker_args[@]}" "$image" "$@"
}

bootstrap_image() {
	run_image bash -c '
		set -euo pipefail
		bun install --frozen-lockfile
		python3 - <<PY
import importlib.metadata as metadata
for name in ["XlsxWriter", "fastxlsx", "fastexcel", "openpyxl", "polars", "psutil", "pyarrow", "pyexcelerate", "pyfastexcel", "pyopenxlsx", "python-calamine", "xlsx2csv"]:
	print(f"{name}=={metadata.version(name)}")
PY
		cargo fetch --manifest-path fixtures/benchmarks/runners/rust-calamine/Cargo.toml
		cargo fetch --manifest-path fixtures/benchmarks/runners/rust-xlsxwriter/Cargo.toml
		cargo build --release --manifest-path fixtures/benchmarks/runners/rust-calamine/Cargo.toml
		cargo build --release --manifest-path fixtures/benchmarks/runners/rust-xlsxwriter/Cargo.toml
		(cd fixtures/benchmarks/runners/excelize && go mod download)
		(cd fixtures/benchmarks/runners/excelize && go build -o /tmp/ascend-excelize-runner .)
		mvn -q -Dmaven.repo.local="$MAVEN_REPO_LOCAL" -f fixtures/benchmarks/runners/apache-poi/pom.xml dependency:go-offline compile
		mvn -q -Dmaven.repo.local="$MAVEN_REPO_LOCAL" -f fixtures/benchmarks/runners/fastexcel-java/pom.xml dependency:go-offline compile
		dotnet restore fixtures/benchmarks/runners/closedxml/ClosedXmlRunner.csproj
		if [[ "${ACCEPT_NPOI_OSMF_LICENSE:-}" == "1" ]]; then
			dotnet restore fixtures/benchmarks/runners/npoi/NpoiRunner.csproj
		fi
		bun --version
		node --version
		python3 --version
		cargo --version
		rustc --version
		go version
		java -version
		mvn --version
		dotnet --info
	'
}

case "$command" in
	build)
		build_image
		;;
	bootstrap)
		shift || true
		bootstrap_image "$@"
		;;
	run)
		shift || true
		run_image "$@"
		;;
	*)
		run_image "$@"
		;;
esac
