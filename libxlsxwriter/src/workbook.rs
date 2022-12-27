use crate::CStringHelper;

use super::{Chart, ChartType, Format, Worksheet, XlsxError};
use std::cell::RefCell;
use std::collections::HashMap;
use std::ffi::CString;
use std::os::raw::c_char;
use std::pin::Pin;
use std::rc::Rc;

/// The Workbook is the main object exposed by the libxlsxwriter library. It represents the entire spreadsheet as you see it in Excel and internally it represents the Excel file as it is written on disk.
///
/// ```rust
/// use xlsxwriter::prelude::*;
/// fn main() -> Result<(), XlsxError> {
///     let workbook = Workbook::new("test-workbook.xlsx")?;
///     let mut worksheet = workbook.add_worksheet(None)?;
///     worksheet.write_string(0, 0, "Hello Excel", None)?;
///     workbook.close()
/// }
/// ```
pub struct Workbook {
    workbook: *mut libxlsxwriter_sys::lxw_workbook,
    pub(crate) const_str: Rc<RefCell<Vec<Pin<Box<CString>>>>>,
    format_map: Rc<RefCell<HashMap<Format, *mut libxlsxwriter_sys::lxw_format>>>,
}

impl Workbook {
    pub(crate) fn get_internal_format(
        &self,
        format: &Format,
    ) -> Result<*mut libxlsxwriter_sys::lxw_format, XlsxError> {
        let mut map = self.format_map.borrow_mut();
        if let Some(p) = map.get(format) {
            Ok(*p)
        } else {
            unsafe {
                let new_format = libxlsxwriter_sys::workbook_add_format(self.workbook);
                format.set_internal_format(new_format)?;
                map.insert(format.clone(), new_format);
                Ok(new_format)
            }
        }
    }

    pub(crate) fn get_internal_option_format(
        &self,
        format: Option<&Format>,
    ) -> Result<*mut libxlsxwriter_sys::lxw_format, XlsxError> {
        if let Some(format) = format {
            Ok(self.get_internal_format(format)?)
        } else {
            Ok(std::ptr::null_mut())
        }
    }

    pub(crate) fn register_str(&self, s: &str) -> Result<*const c_char, XlsxError> {
        let c = Box::pin(CString::new(s)?);
        let p = c.as_ptr();
        self.const_str.borrow_mut().push(c);
        Ok(p)
    }

    pub(crate) fn register_option_str(&self, s: Option<&str>) -> Result<*const c_char, XlsxError> {
        if let Some(s) = s {
            let c = Box::pin(CString::new(s)?);
            let p = c.as_ptr();
            self.const_str.borrow_mut().push(c);
            Ok(p)
        } else {
            Ok(std::ptr::null())
        }
    }

    /// This function is used to create a new Excel workbook with a given filename.
    /// When specifying a filename it is recommended that you use an .xlsx extension or Excel will generate a warning when opening the file.
    pub fn new(filename: &str) -> Result<Workbook, XlsxError> {
        unsafe {
            let workbook_name = Box::pin(CString::new(filename)?);
            let raw_workbook = libxlsxwriter_sys::workbook_new(workbook_name.as_ptr());
            if raw_workbook.is_null() {
                unreachable!()
            }
            Ok(Workbook {
                workbook: raw_workbook,
                const_str: Rc::new(RefCell::new(vec![workbook_name])),
                format_map: Rc::new(RefCell::new(HashMap::new())),
            })
        }
    }

    /// This function is the same as the [`Workbook::new()`] constructor but allows additional options to be set.
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// let workbook = Workbook::new_with_options("test-workbook_with_options.xlsx", true, Some("target"), true)?;
    /// let mut worksheet = workbook.add_worksheet(None)?;
    /// worksheet.write_string(0, 0, "Hello Excel", None)?;
    /// workbook.close()
    /// # }
    /// ```    
    ///
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
    ) -> Result<Workbook, XlsxError> {
        let mut c_string_helper = CStringHelper::new();
        let tmpdir_ptr = c_string_helper.add_opt(tmpdir)?;

        unsafe {
            let mut workbook_options = libxlsxwriter_sys::lxw_workbook_options {
                constant_memory: constant_memory as u8,
                tmpdir: tmpdir_ptr as *mut c_char,
                use_zip64: use_zip64 as u8,
            };

            let workbook_name = Box::pin(CString::new(filename).expect("Null Error"));

            let raw_workbook =
                libxlsxwriter_sys::workbook_new_opt(workbook_name.as_ptr(), &mut workbook_options);
            if raw_workbook.is_null() {
                unreachable!()
            }
            Ok(Workbook {
                workbook: raw_workbook,
                const_str: Rc::new(RefCell::new(vec![workbook_name])),
                format_map: Rc::new(RefCell::new(HashMap::new())),
            })
        }
    }

    pub fn add_worksheet<'a>(
        &'a self,
        sheet_name: Option<&str>,
    ) -> Result<Worksheet<'a>, XlsxError> {
        let name_cstr =
            sheet_name.map(|x| Box::pin(CString::new(x).expect("Cannot create CString")));
        unsafe {
            if let Some(sheet_name) = name_cstr.as_ref() {
                let result = libxlsxwriter_sys::workbook_validate_sheet_name(
                    self.workbook,
                    sheet_name.as_ptr(),
                );
                if result != libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                    return Err(XlsxError::new(result));
                }
            }

            let worksheet = libxlsxwriter_sys::workbook_add_worksheet(
                self.workbook,
                name_cstr
                    .as_ref()
                    .map(|x| x.as_ptr())
                    .unwrap_or(std::ptr::null()),
            );

            if let Some(name) = name_cstr {
                self.const_str.borrow_mut().push(name)
            }

            if worksheet.is_null() {
                return Err(XlsxError::unknown_error());
            }

            Ok(Worksheet {
                _workbook: self,
                worksheet,
            })
        }
    }

    /// This function returns a [`Worksheet`] object reference based on its name.
    pub fn get_worksheet<'a>(
        &'a self,
        sheet_name: &str,
    ) -> Result<Option<Worksheet<'a>>, XlsxError> {
        unsafe {
            let worksheet = libxlsxwriter_sys::workbook_get_worksheet_by_name(
                self.workbook,
                CString::new(sheet_name)?.as_c_str().as_ptr(),
            );
            if worksheet.is_null() {
                Ok(None)
            } else {
                Ok(Some(Worksheet {
                    _workbook: self,
                    worksheet,
                }))
            }
        }
    }

    /// Create new format struct.
    ///
    /// This function available only for compatibility. Please use [`Format::new`] to create new Format object.
    #[deprecated(since = "0.6", note = "Replaced with Format::new()")]
    pub fn add_format(&self) -> Format {
        Format::new()
    }

    /// [`Workbook::add_chart`] function creates a new chart object that can be added to a worksheet.
    /// Available chart types are defined in [`ChartType`].
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
    /// # use xlsxwriter::prelude::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// let workbook = Workbook::new("test-workbook-define_name.xlsx")?;
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

    /// The [`Workbook::close`] function closes a Workbook object, writes the Excel file to disk,
    /// frees any memory allocated internally to the Workbook and frees the object itself.
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
