use super::{error, Chart, ChartType, Format, Worksheet, XlsxError};
use std::cell::RefCell;
use std::ffi::CString;
use std::os::raw::c_char;
use std::rc::Rc;

/// The Workbook is the main object exposed by the libxlsxwriter library. It represents the entire spreadsheet as you see it in Excel and internally it represents the Excel file as it is written on disk.
///
/// ```rust
/// use xlsxwriter::*;
/// fn main() -> Result<(), XlsxError> {
///     let workbook = Workbook::new("test-workbook.xlsx");
///     let mut worksheet = workbook.add_worksheet(None)?;
///     worksheet.write_string(0, 0, "Hello Excel", None)?;
///     workbook.close()
/// }
/// ```
pub struct Workbook {
    workbook: *mut libxlsxwriter_sys::lxw_workbook,
    _workbook_name: CString,
    pub(crate) const_str: Rc<RefCell<Vec<Vec<u8>>>>,
}

impl Workbook {
    /// This function is used to create a new Excel workbook with a given filename.
    /// When specifying a filename it is recommended that you use an .xlsx extension or Excel will generate a warning when opening the file.
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
                const_str: Rc::new(RefCell::new(Vec::new())),
            }
        }
    }

    /// This function is the same as the [`Workbook::new()`] constructor but allows additional options to be set.
    /// The options that can be set are:
    /// * `constant_memory`: This option reduces the amount of data stored in memory so that large files can be written efficiently.
    ///   This option is off by default. See the note below for limitations when this mode is on.
    /// * `tmpdir`: libxlsxwriter stores workbook data in temporary files prior to assembling the final XLSX file. The temporary
    ///   files are created in the system's temp directory. If the default temporary directory isn't accessible to your application,
    ///   or doesn't contain enough space, you can specify an alternative location using the tmpdir option.
    /// * `use_zip64`: Make the zip library use ZIP64 extensions when writing very large xlsx files to allow the zip container, or
    ///   individual XML files within it, to be greater than 4 GB. See ZIP64 on Wikipedia for more information. This option is
    ///   off by default.
    ///
    /// ### Note
    /// In constant_memory mode each row of in-memory data is written to disk and then freed when a new row is started via one
    /// of the `Worksheet::write_*()` functions. Therefore, once this option is active data should be written in sequential row
    /// by row order. For this reason [`Worksheet::merge_range()`] and some other row based functionality doesn't work in this mode.
    /// See [Constant Memory Mode](https://libxlsxwriter.github.io/working_with_memory.html#ww_mem_constant) for more details.
    ///
    /// Also, in `constant_memory` mode the library uses temp file storage for worksheet data. This can lead to an issue on OSes
    /// that map the `/tmp` directory into memory since it is possible to consume the "system" memory even though the "process"
    /// memory remains constant. In these cases you should use an alternative temp file location by using the `tmpdir` option shown
    /// above. See [Constant memory mode and the /tmp directory](https://libxlsxwriter.github.io/working_with_memory.html#ww_mem_temp)
    /// for more details.
    pub fn new_with_options(
        filename: &str,
        constant_memory: bool,
        tmpdir: Option<&str>,
        use_zip64: bool,
    ) -> Workbook {
        let tmpdir_vec = tmpdir.map(|x| CString::new(x).unwrap().as_bytes_with_nul().to_vec());

        unsafe {
            let tmpdir_ptr;
            if let Some(tmpdir) = tmpdir_vec.as_ref() {
                tmpdir_ptr = tmpdir.as_ptr();
            } else {
                tmpdir_ptr = std::ptr::null();
            }

            let mut workbook_options = libxlsxwriter_sys::lxw_workbook_options {
                constant_memory: constant_memory as u8,
                tmpdir: tmpdir_ptr as *mut c_char,
                use_zip64: use_zip64 as u8,
            };

            let workbook_name = CString::new(filename).expect("Null Error");

            let raw_workbook = libxlsxwriter_sys::workbook_new_opt(
                workbook_name.as_c_str().as_ptr(),
                &mut workbook_options,
            );
            if raw_workbook.is_null() {
                unreachable!()
            }
            Workbook {
                workbook: raw_workbook,
                _workbook_name: workbook_name,
                const_str: Rc::new(RefCell::new(Vec::new())),
            }
        }
    }
    pub fn add_worksheet<'a>(
        &'a self,
        sheet_name: Option<&str>,
    ) -> Result<Worksheet<'a>, XlsxError> {
        let name_vec = sheet_name.map(|x| CString::new(x).unwrap().as_bytes_with_nul().to_vec());
        unsafe {
            if let Some(sheet_name) = name_vec.as_ref() {
                let result = libxlsxwriter_sys::workbook_validate_sheet_name(
                    self.workbook,
                    sheet_name.as_ptr() as *const c_char,
                );
                if result != libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                    return Err(XlsxError::new(result));
                }
            }

            let worksheet = libxlsxwriter_sys::workbook_add_worksheet(
                self.workbook,
                name_vec
                    .as_ref()
                    .map(|x| x.as_ptr() as *const c_char)
                    .unwrap_or(std::ptr::null()),
            );

            if let Some(name) = name_vec {
                self.const_str.borrow_mut().push(name);
            }

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

    pub fn add_format(&self) -> Format {
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

    pub fn add_chart(&self, chart_type: ChartType) -> Chart {
        unsafe {
            let chart = libxlsxwriter_sys::workbook_add_chart(self.workbook, chart_type.value());
            if chart.is_null() {
                unreachable!();
            }

            Chart {
                _workbook: self,
                chart,
            }
        }
    }

    /// This function is used to defined a name that can be used to represent a value,
    /// a single cell or a range of cells in a workbook:
    /// These defined names can then be used in formulas:
    /// ```rust
    /// # use xlsxwriter::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// let workbook = Workbook::new("test-workbook-define_name.xlsx");
    /// let mut worksheet = workbook.add_worksheet(None)?;
    /// workbook.define_name("Exchange_rate", "=0.95");
    /// worksheet.write_formula(0, 0, "=Exchange_rate", None);
    /// # workbook.close()
    /// # }
    /// ```
    pub fn define_name(&self, name: &str, formula: &str) -> Result<(), XlsxError> {
        unsafe {
            let result = libxlsxwriter_sys::workbook_define_name(
                self.workbook,
                CString::new(name).expect("Null Error").as_c_str().as_ptr(),
                CString::new(formula)
                    .expect("Null Error")
                    .as_c_str()
                    .as_ptr(),
            );

            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
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
