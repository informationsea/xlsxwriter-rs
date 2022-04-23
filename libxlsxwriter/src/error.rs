use std::error::Error;
use std::ffi;
use std::fmt::{self, Display};

pub const UNKNOWN_ERROR_CODE: libxlsxwriter_sys::lxw_error = 1000;
pub(crate) const NUMBER_OF_COLUMNS_IS_NOT_MATCHED: libxlsxwriter_sys::lxw_error = 1001;

#[derive(Debug)]
pub struct XlsxError {
    pub(crate) error: libxlsxwriter_sys::lxw_error,
}

impl Error for XlsxError {}

impl XlsxError {
    pub fn new(error: libxlsxwriter_sys::lxw_error) -> XlsxError {
        XlsxError { error }
    }
}

impl Display for XlsxError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self.error {
            UNKNOWN_ERROR_CODE => {
                write!(f, "Unknown Error")
            }
            NUMBER_OF_COLUMNS_IS_NOT_MATCHED => {
                write!(
                    f,
                    "Number of columns in an option is not equal to table size"
                )
            }
            _ => unsafe {
                match ffi::CStr::from_ptr(libxlsxwriter_sys::lxw_strerror(self.error)).to_str() {
                    Ok(error_text) => write!(f, "{}", error_text),
                    Err(e) => write!(f, "Cannot get xlsx error text: {}", e),
                }
            },
        }
    }
}
