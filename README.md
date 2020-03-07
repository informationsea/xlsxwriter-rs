xlsxwriter-rs
=============
[![Build Status](https://travis-ci.org/informationsea/xlsxwriter-rs.svg?branch=master)](https://travis-ci.org/informationsea/xlsxwriter-rs)
[![Build status](https://ci.appveyor.com/api/projects/status/lf6j3rvk0p5xd0gn?svg=true)](https://ci.appveyor.com/project/informationsea/xlsxwriter-rs)
![GitHub](https://img.shields.io/github/license/informationsea/xlsxwriter-rs)
![GitHub top language](https://img.shields.io/github/languages/top/informationsea/xlsxwriter-rs)
[![Crates.io](https://img.shields.io/crates/v/xlsxwriter)](https://crates.io/crates/xlsxwriter)
[![Docs.rs](https://docs.rs/xlsxwriter/badge.svg)](https://docs.rs/xlsxwriter)

Rust binding of [libxlsxwriter](https://github.com/jmcnamara/libxlsxwriter)

Supported Features
------------------

* 100% compatible Excel XLSX files.
* Full Excel formatting.
* Merged cells.
* Autofilters.
* Data validation and drop down lists.
* Worksheet PNG/JPEG images.

Coming soon

* Charts.

Example
-------

![Result Image](images/simple1.png)

```rust
let workbook = Workbook::new("target/simple1.xlsx");
let format1 = workbook.add_format().set_font_color(FormatColor::Red);

let format2 = workbook
    .add_format()
    .set_font_color(FormatColor::Blue)
    .set_underline(FormatUnderline::Single);

let format3 = workbook
    .add_format()
    .set_font_color(FormatColor::Green)
    .set_align(FormatAlignment::CenterAcross)
    .set_align(FormatAlignment::VerticalCenter);

let mut sheet1 = workbook.add_worksheet(None)?;
sheet1.write_string(0, 0, "Red text", Some(&format1))?;
sheet1.write_number(0, 1, 20., None)?;
sheet1.write_formula_num(1, 0, "=10+B1", None, 30.)?;
sheet1.write_url(
    1,
    1,
    "https://github.com/informationsea/xlsxwriter-rs",
    Some(&format2),
)?;
sheet1.merge_range(2, 0, 3, 2, "Hello, world", Some(&format3))?;

sheet1.set_selection(1, 0, 1, 2);
sheet1.set_tab_color(FormatColor::Cyan);
workbook.close()?;
```

Run Example with Cargo
-------
You can explore this example locally following the steps below:

1. clone or fork this repository
2. `git submodule update --init --recursive`
    - The above command downloads the required C library so we can compile
3. `cargo run --example hello_spreadsheet`

You should then see the example in `target/simple1.xlsx`. Enjoy!
