xlsxwriter-rs
=============
[![Build Status](https://travis-ci.org/informationsea/xlsxwriter-rs.svg?branch=master)](https://travis-ci.org/informationsea/xlsxwriter-rs)
![GitHub](https://img.shields.io/github/license/informationsea/xlsxwriter-rs)
![GitHub top language](https://img.shields.io/github/languages/top/informationsea/xlsxwriter-rs)

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

![Result Image](images\simple1.png)

```rust
let workbook = Workbook::new("target/simple1.xlsx");
let mut format1 = workbook.get_format();
format1.set_font_color(FormatColor::Red);

let mut format2 = workbook.get_format();
format2
    .set_font_color(FormatColor::Blue)
    .set_underline(FormatUnderline::Single);

let mut format3 = workbook.get_format();
format3
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

