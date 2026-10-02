# syntax=docker/dockerfile:1
# Excelize v2.10.1 oracle helper image for remote Docker runs.
# Build (see ../remote/README.md):
#   docker --context pc build --build-context oracle-remote=../remote \
#     -t excelbench-excelize-oracle:2.10.1 .
FROM golang:1.25-trixie@sha256:2c4c60ef415fbfa5e90300722293bef36c5e63fae17570ce18f580af933dbd73 AS build

ENV GOTOOLCHAIN=local CGO_ENABLED=0
WORKDIR /src
COPY go.mod go.sum ./
RUN go mod download && go mod verify
COPY *.go ./
RUN go test ./... && go build -trimpath -o /out/excelize-oracle .

FROM debian:trixie-slim@sha256:a99cfc517144bc59b1978475ec53b46ecabec7e43635402ee5b77cc54cd1b20a

RUN apt-get update \
    && DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends python3 \
    && rm -rf /var/lib/apt/lists/*

COPY --from=build /out/excelize-oracle /opt/excelbench/excelize/excelize-oracle
COPY --from=oracle-remote container_entry.py /opt/excelbench/remote/container_entry.py

LABEL org.excelbench.adapter=excelize \
      org.excelbench.library=github.com/xuri/excelize/v2 \
      org.excelbench.library-version=v2.10.1

ENTRYPOINT ["python3", "/opt/excelbench/remote/container_entry.py", \
    "--cwd", "/opt/excelbench/excelize", "--", \
    "/opt/excelbench/excelize/excelize-oracle"]
