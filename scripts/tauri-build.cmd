@echo off
call "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\Common7\Tools\VsDevCmd.bat" -arch=x64 -host_arch=x64
set "RUST_TOOLCHAIN_BIN=%USERPROFILE%\.rustup\toolchains\stable-x86_64-pc-windows-msvc\bin"
set "PATH=%RUST_TOOLCHAIN_BIN%;%USERPROFILE%\.cargo\bin;%PATH%"
set "CARGO_TARGET_DIR=C:\cargo-target\clinica"
if not exist "%CARGO_TARGET_DIR%" mkdir "%CARGO_TARGET_DIR%"
npm run tauri:build
