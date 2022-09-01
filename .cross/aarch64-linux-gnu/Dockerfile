FROM debian:oldstable-slim
RUN apt-get update && apt-get install -y clang crossbuild-essential-arm64
ENV SYSROOT /usr/aarch64-linux-gnu
ENV CC_aarch64_unknown_linux_gnu aarch64-linux-gnu-gcc
ENV CXX_aarch64_unknown_linux_gnu aarch64-linux-gnu-g++
ENV CARGO_TARGET_AARCH64_UNKNOWN_LINUX_GNU_LINKER aarch64-linux-gnu-gcc
