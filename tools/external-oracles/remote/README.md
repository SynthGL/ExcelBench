# Remote Docker External Oracles

Runs the write-only `apache-poi` (Java) and `excelize` (Go) ExcelBench adapters
through containers on a remote Docker context, so the benchmark host needs no
local JDK, Go toolchain, or local Docker daemon.

## Build the images

The build context uploads from this checkout to the remote daemon. Each
Dockerfile pulls `container_entry.py` from this directory through a named build
context.

```bash
cd tools/external-oracles/apache-poi
docker --context pc build --build-context oracle-remote=../remote \
  -t excelbench-poi-oracle:5.5.1 .

cd ../excelize
docker --context pc build --build-context oracle-remote=../remote \
  -t excelbench-excelize-oracle:2.11.0 .
```

The POI image runs `build.sh` (which runs `fetch_deps.py`, verifying every
Maven Central jar against its published checksum, then compiles `PoiOracle` and
runs its self-test). The Excelize image runs `go mod verify` and `go test ./...`
against the pinned `go.mod`/`go.sum` before building the helper binary.

## Run the benchmark

```bash
EXCELBENCH_ORACLE_DOCKER_CONTEXT=pc \
  excelbench benchmark -t fixtures/excel -o results-dir -a apache-poi -a excelize
```

When `EXCELBENCH_ORACLE_DOCKER_CONTEXT` is set, `external_oracle_catalog()`
points the `apache-poi` and `excelize` tools at `docker_oracle.py` with that
context and the image tags above. When it is unset, the catalog is unchanged
and uses the local Java and Go helpers.

## Transport

Bind mounts would resolve on the remote host's filesystem, so nothing is
mounted. `docker_oracle.py` reads the oracle request JSON from stdin and:

1. Base64-embeds the contents of `input_path` (or of an existing `output_path`
   when no `input_path` is given, matching the helpers' read fallback) and of
   every local file named by `payload.images[].path`, rewriting those paths to
   container paths under `/work/in/`.
2. Rewrites `output_path` to `/work/out/<same file name>`.
3. Runs `docker --context <ctx> run --rm -i --pull never --network none <image>`
   with the envelope on stdin. `container_entry.py` (the image entrypoint)
   writes the embedded files, runs the helper with the rewritten request, and
   replies with one JSON envelope holding the helper's exact stdout and stderr
   bytes, its return code, and the bytes of the written workbook.
4. Writes the workbook to the original local `output_path`, prints the helper's
   stdout unchanged, forwards its stderr, and exits with the helper's return
   code, so helper errors surface exactly as they do locally.

If Docker itself fails (missing image, unreachable context), the wrapper
prints `{"error": "docker_oracle_transport_failed", ...}` and exits with
Docker's non-zero code. Helper containers have no network access and never
pull images implicitly.

## Runtime versions

Recorded from the images built on `pc` (linux/amd64): POI on 2026-10-01, Excelize 2.11.0 on 2026-10-02 (UTC).

| Image | Image ID | Library | Runtime |
| --- | --- | --- | --- |
| `excelbench-poi-oracle:5.5.1` | `ef46106ed14d` | Apache POI 5.5.1 (`poi`, `poi-ooxml`, `poi-ooxml-lite`), xmlbeans 5.3.0 | Eclipse Temurin OpenJDK 21.0.12.1+1 LTS (`javac 21.0.12.1`), Python 3.14.4, Ubuntu 26.04.1 LTS |
| `excelbench-excelize-oracle:2.11.0` | `79f691f8e6d0` | `github.com/xuri/excelize/v2` v2.11.0 | Go 1.25.14 (static, `CGO_ENABLED=0`), Python 3.13.5, Debian 13 (trixie) |

Base images are pinned by digest in the Dockerfiles:

- `eclipse-temurin:21-jdk@sha256:3e3c176ffed168beb42c607be9bc1639b466cf00261a0fb04425562c9d0c5c2b`
- `golang:1.25-trixie@sha256:2c4c60ef415fbfa5e90300722293bef36c5e63fae17570ce18f580af933dbd73`
- `debian:trixie-slim@sha256:a99cfc517144bc59b1978475ec53b46ecabec7e43635402ee5b77cc54cd1b20a`

The remaining POI runtime jars are the exact set pinned in
`../apache-poi/fetch_deps.py`: commons-compress 1.28.0, commons-io 2.21.0,
commons-lang3 3.18.0, curvesapi 1.08, log4j-api 2.24.3, commons-collections4
4.5.0, commons-codec 1.20.0, commons-math3 3.6.1, SparseBitSet 1.3.
