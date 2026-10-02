# syntax=docker/dockerfile:1
# Apache POI 5.5.1 oracle helper image for remote Docker runs.
# Build (see ../remote/README.md):
#   docker --context pc build --build-context oracle-remote=../remote \
#     -t excelbench-poi-oracle:5.5.1 .
FROM eclipse-temurin:21-jdk@sha256:3e3c176ffed168beb42c607be9bc1639b466cf00261a0fb04425562c9d0c5c2b

RUN apt-get update \
    && DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends python3 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /opt/excelbench/apache-poi
COPY build.sh fetch_deps.py poi_oracle.py ./
COPY src/ src/
# fetch_deps.py verifies Maven Central checksums; build.sh compiles and runs the self-test.
RUN ./build.sh

COPY --from=oracle-remote container_entry.py /opt/excelbench/remote/container_entry.py

LABEL org.excelbench.adapter=apache-poi \
      org.excelbench.library=org.apache.poi:poi-ooxml \
      org.excelbench.library-version=5.5.1

ENTRYPOINT ["python3", "/opt/excelbench/remote/container_entry.py", \
    "--cwd", "/opt/excelbench/apache-poi", "--", \
    "python3", "/opt/excelbench/apache-poi/poi_oracle.py"]
