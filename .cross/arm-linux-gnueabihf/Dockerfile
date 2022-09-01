FROM debian:oldstable-slim
RUN apt-get update && apt-get install -y clang crossbuild-essential-armhf
ENV SYSROOT /usr/arm-linux-gnueabihf
ENV CC_arm_unknown_linux_gnueabihf arm-linux-gnueabihf-gcc
ENV CXX_arm_unknown_linux_gnueabihf arm-linux-gnueabihf-g++
ENV CARGO_TARGET_ARM_UNKNOWN_LINUX_GNUEABIHF_LINKER arm-linux-gnueabihf-gcc
ENV CC_armv7_unknown_linux_gnueabihf arm-linux-gnueabihf-gcc
ENV CXX_armv7_unknown_linux_gnueabihf arm-linux-gnueabihf-g++
ENV CARGO_TARGET_ARMV7_UNKNOWN_LINUX_GNUEABIHF_LINKER arm-linux-gnueabihf-gcc
