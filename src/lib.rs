//! xlsxwriter-rs
//! =============
//!
//! Rust binding of [libxlsxwriter](https://github.com/jmcnamara/libxlsxwriter)
//!
//! ** API of this library is not stable. **
//!
//! Supported Features
//! ------------------
//!
//! * 100% compatible Excel XLSX files.
//! * Full Excel formatting.
//! * Merged cells.
//! * Autofilters.
//! * Data validation and drop down lists.
//! * Worksheet PNG/JPEG images.
//! * Cell comments.
//!
//! Coming soon
//! -----------
//! * Charts.
//!
//! Examples
//! --------
//!
//! ![Result Image](https://github.com/informationsea/xlsxwriter-rs/raw/master/images/simple1.png)
//!
//! ```rust
//! use xlsxwriter::*;
//!
//! # fn main() { let _ = run(); }
//! # fn run() -> Result<(), XlsxError> {
//! let workbook = Workbook::new("simple1.xlsx");
//! let mut format1 = workbook.add_format()
//!     .set_font_color(FormatColor::Red);
//!
//! let mut format2 = workbook.add_format()
//!     .set_font_color(FormatColor::Blue)
//!     .set_underline(FormatUnderline::Single);
//!
//! let mut format3 = workbook.add_format()
//!     .set_font_color(FormatColor::Green)
//!     .set_align(FormatAlignment::CenterAcross)
//!     .set_align(FormatAlignment::VerticalCenter);
//!
//! let mut sheet1 = workbook.add_worksheet(None)?;
//! sheet1.write_string(0, 0, "Red text", Some(&format1))?;
//! sheet1.write_number(0, 1, 20., None)?;
//! sheet1.write_formula_num(1, 0, "=10+B1", None, 30.)?;
//! sheet1.write_url(
//!     1,
//!     1,
//!     "https://github.com/informationsea/xlsxwriter-rs",
//!     Some(&format2),
//! )?;
//! sheet1.merge_range(2, 0, 3, 2, "Hello, world", Some(&format3))?;
//!
//! sheet1.set_selection(1, 0, 1, 2);
//! sheet1.set_tab_color(FormatColor::Cyan);
//! workbook.close()?;
//! # Ok(())
//! # }
//! ```
//!
//! Please read [original libxlsxwriter document](https://libxlsxwriter.github.io/worksheet_8h.html) for description missing functions.
//! Most of this document is based on libxlsxwriter document.

extern crate libxlsxwriter_sys;

mod chart;
mod error;
mod format;
mod validation;
mod workbook;
mod worksheet;

pub use chart::{Chart, ChartDashType, ChartFill, ChartLine, ChartSeries, ChartType};
pub use error::XlsxError;
pub use format::{
    Format, FormatAlignment, FormatBorder, FormatColor, FormatPatterns, FormatScript,
    FormatUnderline,
};
pub use validation::{
    DataValidation, DataValidationCriteria, DataValidationErrorType, DataValidationType,
};
pub use workbook::Workbook;
pub use worksheet::{
    CommentOptions, DateTime, GridLines, HeaderFooterOptions, ImageOptions, PaperType, Protection,
    RowColOptions, Worksheet, WorksheetCol, WorksheetRow,
};

use std::ffi::CString;

fn convert_bool(value: bool) -> u8 {
    let result = if value {
        libxlsxwriter_sys::lxw_boolean_LXW_TRUE
    } else {
        libxlsxwriter_sys::lxw_boolean_LXW_FALSE
    };
    result as u8
}

fn convert_str(value: &str) -> Vec<u8> {
    CString::new(value).unwrap().as_bytes_with_nul().to_vec()
}

#[cfg(test)]
mod test;
