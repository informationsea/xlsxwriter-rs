extern crate bindgen;

use std::env;
use std::fs;
use std::io;
use std::path::PathBuf;

const C_FILES: [&str; 39] = [
    "third_party/libxlsxwriter/third_party/tmpfileplus/tmpfileplus.c",
    "third_party/libxlsxwriter/third_party/minizip/ioapi.c",
    "third_party/libxlsxwriter/third_party/minizip/zip.c",
    "third_party/libxlsxwriter/third_party/md5/md5.c",
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
    "third_party/libxlsxwriter/src/packager.c",
    "third_party/libxlsxwriter/src/relationships.c",
    "third_party/libxlsxwriter/src/shared_strings.c",
    "third_party/libxlsxwriter/src/styles.c",
    "third_party/libxlsxwriter/src/theme.c",
    "third_party/libxlsxwriter/src/utility.c",
    "third_party/libxlsxwriter/src/vml.c",
    "third_party/libxlsxwriter/src/workbook.c",
    "third_party/libxlsxwriter/src/worksheet.c",
    "third_party/libxlsxwriter/src/xmlwriter.c",
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
        .flag_if_supported("-Wno-unused-parameter")
        .include("third_party/zlib");
    for path in &C_FILES[..] {
        assert_file_exists(path)?;
        build.file(path);
    }
    if cfg!(windows) {
        build
            .file("third_party/libxlsxwriter/third_party/minizip/iowin32.c")
            .flag_if_supported("/utf-8")
            .include("include");
    }
    build.compile("libxlsxwriter.a");

    // The bindgen::Builder is the main entry point
    // to bindgen, and lets you build up options for
    // the resulting bindings.
    let bindings = bindgen::Builder::default()
        .generate_comments(false)
        .clang_arg("-Iinclude")
        .header("wrapper.h")
        .whitelist_function("^(chart|chartsheet|workbook|worksheet|format|lxw)_.*")
        .whitelist_type("^lxw_.*")
        .whitelist_var("^lxw_.*")
        .blacklist_function("_get_image_properties")
        .generate()
        .expect("Unable to generate bindings");

    // Write the bindings to the $OUT_DIR/bindings.rs file.
    let out_path = PathBuf::from(env::var("OUT_DIR").unwrap());
    bindings
        .write_to_file(out_path.join("bindings.rs"))
        .expect("Couldn't write bindings!");

    Ok(())
}
