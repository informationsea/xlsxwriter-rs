extern crate bindgen;

use std::env;
use std::fs;
use std::io;
use std::path::PathBuf;

const C_FILES: [&str; 26] = [
    "third_party/libxlsxwriter/third_party/tmpfileplus/tmpfileplus.c",
    "third_party/libxlsxwriter/third_party/minizip/ioapi.c",
    "third_party/libxlsxwriter/third_party/minizip/zip.c",
    "third_party/libxlsxwriter/third_party/dtoa/emyg_dtoa.c",
    "third_party/libxlsxwriter/src/app.c",
    "third_party/libxlsxwriter/src/chart.c",
    "third_party/libxlsxwriter/src/chartsheet.c",
    "third_party/libxlsxwriter/src/comment.c",
    "third_party/libxlsxwriter/src/content_types.c",
    "third_party/libxlsxwriter/src/core.c",
    "third_party/libxlsxwriter/src/custom.c",
    "third_party/libxlsxwriter/src/drawing.c",
    "third_party/libxlsxwriter/src/format.c",
    "third_party/libxlsxwriter/src/hash_table.c",
    "third_party/libxlsxwriter/src/metadata.c",
    "third_party/libxlsxwriter/src/packager.c",
    "third_party/libxlsxwriter/src/relationships.c",
    "third_party/libxlsxwriter/src/shared_strings.c",
    "third_party/libxlsxwriter/src/styles.c",
    "third_party/libxlsxwriter/src/table.c",
    "third_party/libxlsxwriter/src/theme.c",
    "third_party/libxlsxwriter/src/utility.c",
    "third_party/libxlsxwriter/src/vml.c",
    "third_party/libxlsxwriter/src/workbook.c",
    "third_party/libxlsxwriter/src/worksheet.c",
    "third_party/libxlsxwriter/src/xmlwriter.c",
];

const ZLIB_FILES: [&str; 15] = [
    "third_party/zlib/adler32.c",
    "third_party/zlib/compress.c",
    "third_party/zlib/crc32.c",
    "third_party/zlib/deflate.c",
    "third_party/zlib/gzclose.c",
    "third_party/zlib/gzlib.c",
    "third_party/zlib/gzread.c",
    "third_party/zlib/gzwrite.c",
    "third_party/zlib/infback.c",
    "third_party/zlib/inffast.c",
    "third_party/zlib/inflate.c",
    "third_party/zlib/inftrees.c",
    "third_party/zlib/trees.c",
    "third_party/zlib/uncompr.c",
    "third_party/zlib/zutil.c",
];

fn assert_file_exists(path: &str) -> io::Result<()> {
    match fs::metadata(path) {
        Ok(_) => Ok(()),
        Err(ref e) if e.kind() == io::ErrorKind::NotFound => {
            panic!(
                "Can't access {}. Did you forget to fetch git submodules?",
                path
            );
        }
        Err(e) => Err(e),
    }
}

fn main() -> io::Result<()> {
    let mut build = cc::Build::new();
    build
        .include("third_party/libxlsxwriter/include")
        .flag_if_supported("-Wno-implicit-function-declaration")
        .flag_if_supported("-Wno-unused-parameter");
    for path in &C_FILES[..] {
        assert_file_exists(path)?;
        build.file(path);
    }

    if env::var("CARGO_FEATURE_SYSTEM_ZLIB").is_ok() {
        println!("cargo:rustc-link-lib=z");
    } else {
        build.include("third_party/zlib");
        for path in &ZLIB_FILES[..] {
            assert_file_exists(path)?;
            build.file(path);
        }
    }

    if env::var("CARGO_FEATURE_NO_MD5").is_ok() {
        build.define("USE_NO_MD5", None);
    } else if env::var("CARGO_FEATURE_USE_OPENSSL_MD5").is_ok() {
        build.define("USE_OPENSSL_MD5", None);
        println!("cargo:rustc-link-lib=crypto");
    } else {
        build.file("third_party/libxlsxwriter/third_party/md5/md5.c");
    }

    if cfg!(windows) {
        build
            .file("third_party/libxlsxwriter/third_party/minizip/iowin32.c")
            .flag_if_supported("/utf-8")
            .include("include");
    }

    // Make `libxlsxwriter` use DTOA for number formating to avoid locale-specific C functions.
    build.define("USE_DTOA_LIBRARY", None);

    build.compile("libxlsxwriter.a");

    // The bindgen::Builder is the main entry point
    // to bindgen, and lets you build up options for
    // the resulting bindings.
    let builder = bindgen::Builder::default()
        .generate_comments(false)
        .clang_arg("-Iinclude");
    let builder = if let Ok(sysroot) = env::var("SYSROOT") {
        builder.clang_arg(format!("--sysroot={}", sysroot))
    } else {
        builder
    };

    let bindings = builder
        .header("wrapper.h")
        .allowlist_function("^(chart|chartsheet|workbook|worksheet|format|lxw)_.*")
        .allowlist_type("^lxw_.*")
        .allowlist_var("^lxw_.*")
        .blocklist_function("_get_image_properties")
        .generate()
        .expect("Unable to generate bindings");

    // Write the bindings to the $OUT_DIR/bindings.rs file.
    let out_path = PathBuf::from(env::var("OUT_DIR").unwrap());
    bindings
        .write_to_file(out_path.join("bindings.rs"))
        .expect("Couldn't write bindings!");

    Ok(())
}
