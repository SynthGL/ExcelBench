#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

JAVA_BIN="${JAVA_BIN:-}"
JAVAC_BIN="${JAVAC_BIN:-}"

if [ -z "$JAVA_BIN" ] && [ -x /opt/homebrew/opt/openjdk/bin/java ]; then
  JAVA_BIN=/opt/homebrew/opt/openjdk/bin/java
fi
if [ -z "$JAVA_BIN" ]; then
  JAVA_BIN=java
fi
if ! "$JAVA_BIN" -version >/dev/null 2>&1; then
  echo "usable java not found" >&2
  exit 1
fi

if [ -z "$JAVAC_BIN" ] && [ -x /opt/homebrew/opt/openjdk/bin/javac ]; then
  JAVAC_BIN=/opt/homebrew/opt/openjdk/bin/javac
fi
if [ -z "$JAVAC_BIN" ]; then
  JAVAC_BIN=javac
fi
if ! "$JAVAC_BIN" -version >/dev/null 2>&1; then
  echo "usable javac not found" >&2
  exit 1
fi

python3 fetch_deps.py
mkdir -p build/classes
"$JAVAC_BIN" -cp "deps/lib/*" -d build/classes src/PoiOracle.java
"$JAVA_BIN" -cp "build/classes:deps/lib/*" PoiOracle --self-test
