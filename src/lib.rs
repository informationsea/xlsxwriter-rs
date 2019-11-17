extern crate libxlsxwriter_sys;

mod error;
mod format;
mod validation;
mod workbook;
mod worksheet;

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
    DateTime, GridLines, HeaderFooterOptions, ImageOptions, PaperType, Protection, RowColOptions,
    Worksheet,
};

fn convert_bool(value: bool) -> u8 {
    let result = if value {
        libxlsxwriter_sys::lxw_boolean_LXW_TRUE
    } else {
        libxlsxwriter_sys::lxw_boolean_LXW_FALSE
    };
    result as u8
}

#[cfg(test)]
mod test;
