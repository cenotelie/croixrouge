ARG BUILD_FLAGS=
ARG BUILD_TARGET=debug

## Base image with common dependencies
FROM buildpack-deps:24.04-curl AS base
LABEL maintainer="Laurent Wouters <lwouters@cenotelie.fr>" vendor="Cénotélie"  description="Croix-Rouge"
# add packages
RUN apt-get update && apt-get install -y --no-install-recommends \
		build-essential \
		pkg-config \
		git \
		libc6-dev \
		libfontconfig1 \
	&& rm -rf /var/lib/apt/lists/*

# add custom user
RUN groupmod -n croixrouge ubuntu && usermod -l croixrouge -d /home/croixrouge ubuntu && mv /home/ubuntu /home/croixrouge
ENV HOME=/home/croixrouge
USER croixrouge



## Builder to build the application
FROM base AS builder
ARG BUILD_FLAGS
# add .Net support
USER root
RUN apt-get update && apt-get install -y --no-install-recommends \
		dotnet-sdk-8.0 \
	&& rm -rf /var/lib/apt/lists/*
USER croixrouge
# add Rust support
ENV PATH="/home/croixrouge/.cargo/bin:${PATH}"
RUN curl https://sh.rustup.rs -sSf | sh -s -- -y \
	&& rm -rf /home/croixrouge/.cargo/registry \
	&& mkdir /home/croixrouge/.cargo/registry
# copy the sources and build
COPY --chown=croixrouge . /home/croixrouge/src
RUN cd /home/croixrouge/src/DistributionCR && dotnet build --runtime linux-x64 --configuration Release
RUN cd /home/croixrouge/src/server && cargo build ${BUILD_FLAGS}



## Final target from the base with the application's binary
FROM base
ARG BUILD_TARGET
# add .Net support
USER root
RUN apt-get update && apt-get install -y --no-install-recommends \
		dotnet-runtime-8.0 \
	&& rm -rf /var/lib/apt/lists/*
USER croixrouge
COPY --from=builder /home/croixrouge/src/DistributionCR/bin/Release/net8.0/linux-x64 /home/croixrouge/payload
COPY --from=builder /home/croixrouge/src/server/target/${BUILD_TARGET}/croixrouge /
ENTRYPOINT ["/croixrouge"]
