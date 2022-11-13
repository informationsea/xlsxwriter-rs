use std::error::Error;
use std::ffi;
use std::fmt::{self, Display};

#[derive(Debug, Clone, PartialEq)]
pub(crate) enum XlsxErrorSource {
    LibXlsxWriter(libxlsxwriter_sys::lxw_error),
    NumberOfColumnsIsNotMatched,
    Unknown,
    NulError(std::ffi::NulError),
}

#[derive(Debug)]
pub struct XlsxError {
    pub(crate) source: XlsxErrorSource,
}

impl Error for XlsxError {}

impl XlsxError {
    pub fn new(error: libxlsxwriter_sys::lxw_error) -> XlsxError {
        XlsxError {
            source: XlsxErrorSource::LibXlsxWriter(error),
        }
    }

    pub fn unknown_error() -> XlsxError {
        XlsxError {
            source: XlsxErrorSource::Unknown,
        }
    }
}

impl Display for XlsxError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.source {
            XlsxErrorSource::Unknown => {
                write!(f, "Unknown Error")
            }
            XlsxErrorSource::NumberOfColumnsIsNotMatched => {
                write!(
                    f,
                    "Number of columns in an option is not equal to table size"
                )
            }
            XlsxErrorSource::NulError(e) => {
                write!(f, "Null bytes in string: {}", e)
            }
            XlsxErrorSource::LibXlsxWriter(error) => unsafe {
                match ffi::CStr::from_ptr(libxlsxwriter_sys::lxw_strerror(*error)).to_str() {
                    Ok(error_text) => write!(f, "{}", error_text),
                    Err(e) => write!(f, "Cannot get xlsx error text: {}", e),
                }
            },
        }
    }
}

impl From<std::ffi::NulError> for XlsxError {
    fn from(e: std::ffi::NulError) -> Self {
        XlsxError {
            source: XlsxErrorSource::NulError(e),
        }
    }
}
