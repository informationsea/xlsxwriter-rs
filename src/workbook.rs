use super::{error, Format, Worksheet, XlsxError};
use std::ffi::CString;

pub struct Workbook {
    workbook: *mut libxlsxwriter_sys::lxw_workbook,
    _workbook_name: CString,
}

impl Workbook {
    pub fn new(filename: &str) -> Workbook {
        unsafe {
            let workbook_name = CString::new(filename).expect("Null Error");
            let raw_workbook = libxlsxwriter_sys::workbook_new(workbook_name.as_c_str().as_ptr());
            if raw_workbook.is_null() {
                unreachable!()
            }
            Workbook {
                workbook: raw_workbook,
                _workbook_name: workbook_name,
            }
        }
    }
    pub fn add_worksheet<'a>(
        &'a self,
        sheet_name: Option<&str>,
    ) -> Result<Worksheet<'a>, XlsxError> {
        unsafe {
            if let Some(sheet_name) = sheet_name {
                let result = libxlsxwriter_sys::workbook_validate_sheet_name(
                    self.workbook,
                    CString::new(sheet_name)
                        .expect("Null Error")
                        .as_c_str()
                        .as_ptr(),
                );
                if result != libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                    return Err(XlsxError::new(result));
                }
            }

            let worksheet = libxlsxwriter_sys::workbook_add_worksheet(
                self.workbook,
                sheet_name
                    .map(|x| CString::new(x).expect("Null Error").as_c_str().as_ptr())
                    .unwrap_or(std::ptr::null()),
            );

            if worksheet.is_null() {
                return Err(XlsxError::new(error::UNKNOWN_ERROR_CODE));
            }

            Ok(Worksheet {
                _workbook: self,
                worksheet,
            })
        }
    }

    pub fn get_worksheet<'a>(&'a self, sheet_name: &str) -> Option<Worksheet<'a>> {
        unsafe {
            let worksheet = libxlsxwriter_sys::workbook_get_worksheet_by_name(
                self.workbook,
                CString::new(sheet_name)
                    .expect("Null Error")
                    .as_c_str()
                    .as_ptr(),
            );
            if worksheet.is_null() {
                None
            } else {
                Some(Worksheet {
                    _workbook: self,
                    worksheet,
                })
            }
        }
    }

    pub fn get_format(&self) -> Format {
        unsafe {
            let format = libxlsxwriter_sys::workbook_add_format(self.workbook);
            if format.is_null() {
                unreachable!();
            }

            Format {
                _workbook: self,
                format,
            }
        }
    }

    pub fn close(mut self) -> Result<(), XlsxError> {
        unsafe {
            let result = libxlsxwriter_sys::workbook_close(self.workbook);
            self.workbook = std::ptr::null_mut();
            match result {
                libxlsxwriter_sys::lxw_error_LXW_NO_ERROR => Ok(()),
                _ => Err(XlsxError::new(result)),
            }
        }
    }
}

impl Drop for Workbook {
    fn drop(&mut self) {
        unsafe {
            if !self.workbook.is_null() {
                libxlsxwriter_sys::workbook_close(self.workbook);
            }
        }
    }
}
