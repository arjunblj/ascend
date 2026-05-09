#!/usr/bin/env bash
set -euo pipefail

tmp_root="${TMPDIR:-/tmp}"
if [[ -z "${JAVA_HOME:-}" && -d /opt/homebrew/opt/openjdk ]]; then
	export JAVA_HOME=/opt/homebrew/opt/openjdk
fi
if [[ -n "${JAVA_HOME:-}" ]]; then
	export PATH="$JAVA_HOME/bin:$PATH"
fi
export MAVEN_REPO_LOCAL="${MAVEN_REPO_LOCAL:-$tmp_root/ascend-m2-repository}"

exec mvn -q -Dmaven.repo.local="$MAVEN_REPO_LOCAL" \
	-f fixtures/benchmarks/runners/fastexcel-java/pom.xml \
	compile org.codehaus.mojo:exec-maven-plugin:3.6.2:java -Dexec.args="$*"
