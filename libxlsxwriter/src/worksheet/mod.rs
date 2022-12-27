pub mod conditional_format;
mod datetime;
pub mod filter;
pub mod table;
pub mod validation;

use crate::CStringHelper;

use super::{convert_bool, Chart, Format, FormatColor, Workbook, XlsxError};
use std::ffi::CString;
use std::os::raw::c_char;

pub use datetime::*;

/// Integer data type to represent a column value. Equivalent to `u16`.
///
/// The maximum column in Excel is 16,384.
pub type WorksheetCol = libxlsxwriter_sys::lxw_col_t;

/// Integer data type to represent a row value. Equivalent to `u32`.
///
/// The maximum row in Excel is 1,048,576.
pub type WorksheetRow = libxlsxwriter_sys::lxw_row_t;

#[derive(Debug, Clone, PartialEq, PartialOrd, Default)]
pub struct DateTime {
    pub year: i16,
    pub month: i8,
    pub day: i8,
    pub hour: i8,
    pub min: i8,
    pub second: f64,
}

/// Options for modifying images inserted via [Worksheet.insert_image_opt()](struct.Worksheet.html#method.insert_image_opt).
#[derive(Debug, Clone, PartialEq, PartialOrd)]
pub struct ImageOptions {
    /// Offset from the left of the cell in pixels.
    pub x_offset: i32,
    /// Offset from the top of the cell in pixels.
    pub y_offset: i32,
    /// X scale of the image as a decimal.
    pub x_scale: f64,
    /// Y scale of the image as a decimal.
    pub y_scale: f64,
}

impl From<&ImageOptions> for libxlsxwriter_sys::lxw_image_options {
    fn from(options: &ImageOptions) -> Self {
        libxlsxwriter_sys::lxw_image_options {
            x_offset: options.x_offset,
            y_offset: options.y_offset,
            x_scale: options.x_scale,
            y_scale: options.y_scale,
            description: std::ptr::null_mut(),
            url: std::ptr::null_mut(),
            tip: std::ptr::null_mut(),
            object_position: 0,
            decorative: 0,
        }
    }
}

/// Paper sizes
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub enum PaperType {
    PrinterDefault,
    Letter,
    Tabloid,
    Ledger,
    Legal,
    Statement,
    Executive,
    A3,
    A4,
    A5,
    B4,
    B5,
    Folio,
    Quarto,
    Other(u8),
}

impl PaperType {
    fn value(self) -> u8 {
        let value = match self {
            PaperType::PrinterDefault => 0,
            PaperType::Letter => 1,
            PaperType::Tabloid => 3,
            PaperType::Ledger => 4,
            PaperType::Legal => 5,
            PaperType::Statement => 6,
            PaperType::Executive => 7,
            PaperType::A3 => 8,
            PaperType::A4 => 9,
            PaperType::A5 => 11,
            PaperType::B4 => 12,
            PaperType::B5 => 13,
            PaperType::Folio => 14,
            PaperType::Quarto => 15,
            PaperType::Other(x) => x.into(),
        };
        value as u8
    }
}

/// Options for header and footer
#[derive(Debug, Clone, PartialEq, PartialOrd)]
pub struct HeaderFooterOptions {
    pub margin: f64,
}

impl From<&HeaderFooterOptions> for libxlsxwriter_sys::lxw_header_footer_options {
    fn from(options: &HeaderFooterOptions) -> libxlsxwriter_sys::lxw_header_footer_options {
        libxlsxwriter_sys::lxw_header_footer_options {
            margin: options.margin,
            image_left: std::ptr::null_mut(),
            image_center: std::ptr::null_mut(),
            image_right: std::ptr::null_mut(),
        }
    }
}

#[derive(Debug, Copy, Clone, PartialEq, PartialOrd)]
pub enum GridLines {
    HideAllGridLines,
    ShowScreenGridLines,
    ShowPrintGridLines,
    ShowAllGridLines,
}

impl GridLines {
    fn value(self) -> u8 {
        let value = match self {
            GridLines::HideAllGridLines => libxlsxwriter_sys::lxw_gridlines_LXW_HIDE_ALL_GRIDLINES,
            GridLines::ShowScreenGridLines => {
                libxlsxwriter_sys::lxw_gridlines_LXW_SHOW_SCREEN_GRIDLINES
            }
            GridLines::ShowPrintGridLines => {
                libxlsxwriter_sys::lxw_gridlines_LXW_SHOW_PRINT_GRIDLINES
            }
            GridLines::ShowAllGridLines => libxlsxwriter_sys::lxw_gridlines_LXW_SHOW_ALL_GRIDLINES,
        };
        value as u8
    }
}

/// Sheet protection
#[derive(Debug, Copy, Clone, PartialEq, PartialOrd)]
pub struct Protection {
    pub no_select_locked_cells: bool,
    pub no_select_unlocked_cells: bool,
    pub format_cells: bool,
    pub format_columns: bool,
    pub format_rows: bool,
    pub insert_columns: bool,
    pub insert_rows: bool,
    pub insert_hyperlinks: bool,
    pub delete_columns: bool,
    pub delete_rows: bool,
    pub sort: bool,
    pub autofilter: bool,
    pub pivot_tables: bool,
    pub scenarios: bool,
    pub objects: bool,
    pub no_content: bool,
    pub no_objects: bool,
}

impl Protection {
    pub fn new() -> Protection {
        Protection {
            no_select_locked_cells: true,
            no_select_unlocked_cells: true,
            format_cells: false,
            format_columns: false,
            format_rows: false,
            insert_columns: false,
            insert_rows: false,
            insert_hyperlinks: false,
            delete_columns: false,
            delete_rows: false,
            sort: false,
            autofilter: false,
            pivot_tables: false,
            scenarios: false,
            objects: false,
            no_content: false,
            no_objects: false,
        }
    }
}

impl Default for Protection {
    fn default() -> Self {
        Protection::new()
    }
}

impl From<&Protection> for libxlsxwriter_sys::lxw_protection {
    fn from(protection: &Protection) -> libxlsxwriter_sys::lxw_protection {
        libxlsxwriter_sys::lxw_protection {
            no_select_locked_cells: convert_bool(protection.no_select_locked_cells),
            no_select_unlocked_cells: convert_bool(protection.no_select_unlocked_cells),
            format_cells: convert_bool(protection.format_cells),
            format_columns: convert_bool(protection.format_columns),
            format_rows: convert_bool(protection.format_rows),
            insert_columns: convert_bool(protection.insert_columns),
            insert_rows: convert_bool(protection.insert_rows),
            insert_hyperlinks: convert_bool(protection.insert_hyperlinks),
            delete_columns: convert_bool(protection.delete_columns),
            delete_rows: convert_bool(protection.delete_rows),
            sort: convert_bool(protection.sort),
            autofilter: convert_bool(protection.autofilter),
            pivot_tables: convert_bool(protection.pivot_tables),
            scenarios: convert_bool(protection.scenarios),
            objects: convert_bool(protection.objects),
            no_content: convert_bool(protection.no_content),
            no_objects: convert_bool(protection.no_objects),
        }
    }
}

/// Options struct for the `set_column()` and `set_row()` functions.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct RowColOptions {
    hidden: bool,
    level: u8,
    collapsed: bool,
}

impl RowColOptions {
    pub fn new(hidden: bool, level: u8, collapsed: bool) -> Self {
        RowColOptions {
            hidden,
            level,
            collapsed,
        }
    }

    pub(crate) fn into_internal(&self) -> libxlsxwriter_sys::lxw_row_col_options {
        libxlsxwriter_sys::lxw_row_col_options {
            hidden: convert_bool(self.hidden),
            level: self.level,
            collapsed: convert_bool(self.collapsed),
        }
    }
}

impl Default for RowColOptions {
    fn default() -> Self {
        RowColOptions {
            hidden: false,
            level: 0,
            collapsed: false,
        }
    }
}

/// Comment display type
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub enum CommentDisplayType {
    Default,
    Hidden,
    Visible,
}

impl Default for CommentDisplayType {
    fn default() -> Self {
        CommentDisplayType::Default
    }
}

impl CommentDisplayType {
    pub(crate) fn into_internal(self) -> libxlsxwriter_sys::lxw_comment_display_types {
        match self {
            CommentDisplayType::Default => {
                libxlsxwriter_sys::lxw_comment_display_types_LXW_COMMENT_DISPLAY_DEFAULT
            }
            CommentDisplayType::Hidden => {
                libxlsxwriter_sys::lxw_comment_display_types_LXW_COMMENT_DISPLAY_HIDDEN
            }
            CommentDisplayType::Visible => {
                libxlsxwriter_sys::lxw_comment_display_types_LXW_COMMENT_DISPLAY_VISIBLE
            }
        }
    }
}

/// Options for modifying comments inserted via `write_comment_opt()`
#[derive(Debug, Clone, PartialEq, PartialOrd)]
pub struct CommentOptions {
    visible: CommentDisplayType,
    author: Option<String>,
    width: Option<u16>,
    height: Option<u16>,
    x_scale: Option<f64>,
    y_scale: Option<f64>,
    color: FormatColor,
    font_name: Option<String>,
    font_size: Option<f64>,
    font_family: Option<u8>,
    start_row: WorksheetRow,
    start_col: WorksheetCol,
    x_offset: i32,
    y_offset: i32,
}

impl CommentOptions {
    pub(crate) fn into_internal(
        &self,
        workbook: &Workbook,
    ) -> Result<libxlsxwriter_sys::lxw_comment_options, XlsxError> {
        Ok(libxlsxwriter_sys::lxw_comment_options {
            visible: self.visible.into_internal() as u8,
            author: workbook.register_option_str(self.author.as_deref())? as *mut c_char,
            width: self.width.unwrap_or_default(),
            height: self.height.unwrap_or_default(),
            x_scale: self.x_scale.unwrap_or_default(),
            y_scale: self.y_scale.unwrap_or_default(),
            color: self.color.value(),
            font_name: workbook.register_option_str(self.font_name.as_deref())? as *mut c_char,
            font_size: self.font_size.unwrap_or_default(),
            font_family: self.font_family.unwrap_or_default(),
            start_row: self.start_row,
            start_col: self.start_col,
            x_offset: self.x_offset,
            y_offset: self.y_offset,
        })
    }
}

/// Default Excel row height in character units.
pub const LXW_DEF_ROW_HEIGHT: f64 = 15.0;
/// Default Excel row height in pixels.
pub const LXW_DEF_ROW_HEIGHT_PIXELS: u32 = 20;
/// Default Excel column width in character units.
pub const LXW_DEF_COL_WIDTH: f64 = 8.43;
/// Default Excel column width in pixels.
pub const LXW_DEF_COL_WIDTH_PIXELS: u32 = 64;

/// The Worksheet object represents an Excel worksheet. It handles operations such as writing data to cells or formatting worksheet layout.
///
/// A Worksheet object isn't created directly. Instead a worksheet is created by calling the `workbook.add_worksheet()` function from a [Workbook](struct.Workbook.html) object:
/// ```rust
/// use xlsxwriter::prelude::*;
/// # fn main() -> Result<(), XlsxError> {
/// let workbook = Workbook::new("test-worksheet.xlsx")?;
/// let mut worksheet = workbook.add_worksheet(None)?;
/// worksheet.write_string(0, 0, "Hello, excel", None)?;
/// workbook.close()
/// # }
/// ```
/// Please read [original libxlsxwriter document](https://libxlsxwriter.github.io/worksheet_8h.html) for description missing functions.
/// Most of this document is based on libxlsxwriter document.
pub struct Worksheet<'a> {
    pub(crate) _workbook: &'a Workbook,
    pub(crate) worksheet: *mut libxlsxwriter_sys::lxw_worksheet,
}

impl<'a> Worksheet<'a> {
    /// This function writes the comment of a cell
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-worksheet_write_comment-1.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// worksheet.write_comment(0, 0, "This is some comment text")?;
    /// worksheet.write_comment(1, 0, "This cell also has a comment")?;
    /// # workbook.close()
    /// # }
    /// ```
    pub fn write_comment(
        &mut self,
        row: WorksheetRow,
        col: WorksheetCol,
        text: &str,
    ) -> Result<(), XlsxError> {
        unsafe {
            let result = libxlsxwriter_sys::worksheet_write_comment(
                self.worksheet,
                row,
                col,
                CString::new(text).unwrap().as_c_str().as_ptr(),
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    pub fn write_comment_opt(
        &mut self,
        row: WorksheetRow,
        col: WorksheetCol,
        text: &str,
        options: &CommentOptions,
    ) -> Result<(), XlsxError> {
        let mut options = options.into_internal(self._workbook)?;
        unsafe {
            let result = libxlsxwriter_sys::worksheet_write_comment_opt(
                self.worksheet,
                row,
                col,
                self._workbook.register_str(text)?,
                &mut options,
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    /// This function writes numeric types to the cell specified by row and column:
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-worksheet_write_number-1.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// worksheet.write_number(0, 0, 123456.0, None)?;
    /// worksheet.write_number(1, 0, 2.3451, None)?;
    /// # workbook.close()
    /// # }
    /// ```
    /// ![Result Image](https://github.com/informationsea/xlsxwriter-rs/raw/master/images/test-worksheet-write_number-1.png)
    ///
    /// The native data type for all numbers in Excel is a IEEE-754 64-bit double-precision floating point, which is also the default type used by worksheet_write_number.
    ///
    /// The format parameter is used to apply formatting to the cell. This parameter can be `None` to indicate no formatting or it can be a Format object.
    ///
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-worksheet_write_number-2.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// worksheet.write_number(0, 0, 1234.567, Some(&Format::new().set_num_format("$#,##0.00")))?;
    /// # workbook.close()
    /// # }
    /// ```
    /// ![Result Image](https://github.com/informationsea/xlsxwriter-rs/raw/master/images/test-worksheet-write_number-2.png)
    ///
    /// ### Note
    /// Excel doesn't support NaN, Inf or -Inf as a number value. If you are writing data that contains these values then your application should convert them to a string or handle them in some other way.
    pub fn write_number(
        &mut self,
        row: WorksheetRow,
        col: WorksheetCol,
        number: f64,
        format: Option<&Format>,
    ) -> Result<(), XlsxError> {
        unsafe {
            let result = libxlsxwriter_sys::worksheet_write_number(
                self.worksheet,
                row,
                col,
                number,
                self._workbook.get_internal_option_format(format)?,
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    /// This function writes a string to the cell specified by row and column:
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-worksheet_write_string-1.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// worksheet.write_string(0, 0, "This phrase is English!", None)?;
    /// # workbook.close()
    /// # }
    /// ```
    /// ![Result Image](https://github.com/informationsea/xlsxwriter-rs/raw/master/images/test-worksheet-write_string-1.png)
    ///
    /// The format parameter is used to apply formatting to the cell. This parameter can be `None` to indicate no formatting or it can be a Format object:
    ///
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-worksheet_write_string-2.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// worksheet.write_string(0, 0, "This phrase is Bold!", Some(&Format::new().set_bold()))?;
    /// # workbook.close()
    /// # }
    /// ```
    /// ![Result Image](https://github.com/informationsea/xlsxwriter-rs/raw/master/images/test-worksheet-write_string-2.png)
    ///
    /// Unicode strings are supported in UTF-8 encoding.
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-worksheet_write_string-3.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// worksheet.write_string(0, 0, "こんにちは、世界！", None)?;
    /// # workbook.close()
    /// # }
    /// ```
    /// ![Result Image](https://github.com/informationsea/xlsxwriter-rs/raw/master/images/test-worksheet-write_string-3.png)
    pub fn write_string(
        &mut self,
        row: WorksheetRow,
        col: WorksheetCol,
        text: &str,
        format: Option<&Format>,
    ) -> Result<(), XlsxError> {
        let mut c_string_helper = CStringHelper::new();
        unsafe {
            let result = libxlsxwriter_sys::worksheet_write_string(
                self.worksheet,
                row,
                col,
                c_string_helper.add(text)?,
                self._workbook.get_internal_option_format(format)?,
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    /// This function writes a formula or function to the cell specified by row and column:
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-worksheet_write_formula-1.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// worksheet.write_formula(0, 0, "=B3 + 6", None)?;
    /// worksheet.write_formula(1, 0, "=SIN(PI()/4)", None)?;
    /// worksheet.write_formula(2, 0, "=SUM(A1:A2)", None)?;
    /// worksheet.write_formula(3, 0, "=IF(A3>1,\"Yes\", \"No\")", None)?;
    /// worksheet.write_formula(4, 0, "=AVERAGE(1, 2, 3, 4)", None)?;
    /// worksheet.write_formula(5, 0, "=DATEVALUE(\"1-Jan-2013\")", None)?;
    /// # workbook.close()
    /// # }
    /// ```
    /// ![Result Image](https://github.com/informationsea/xlsxwriter-rs/raw/master/images/test-worksheet-write_formula-1.png)
    ///
    /// The `format` parameter is used to apply formatting to the cell. This parameter can be `None` to indicate no formatting or it can be a Format object.
    ///
    /// Libxlsxwriter doesn't calculate the value of a formula and instead stores a default value of `0`. The correct formula result is displayed in Excel, as shown in the example above, since it recalculates the formulas when it loads the file. For cases where this is an issue see the `write_formula_num()` function and the discussion in that section.
    ///
    /// Formulas must be written with the US style separator/range operator which is a comma (not semi-colon). Therefore a formula with multiple values should be written as follows:
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-worksheet_write_formula-2.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// // OK
    /// worksheet.write_formula(0, 0, "=SUM(1, 2, 3)", None)?;
    /// // NO. Error on load.
    /// worksheet.write_formula(1, 0, "=SUM(1; 2; 3)", None)?;
    /// # workbook.close()
    /// # }
    /// ```
    /// See also [Working with Formulas](https://libxlsxwriter.github.io/working_with_formulas.html).
    pub fn write_formula(
        &mut self,
        row: WorksheetRow,
        col: WorksheetCol,
        formula: &str,
        format: Option<&Format>,
    ) -> Result<(), XlsxError> {
        let mut c_string_helper = CStringHelper::new();
        unsafe {
            let result = libxlsxwriter_sys::worksheet_write_formula(
                self.worksheet,
                row,
                col,
                c_string_helper.add(formula)?,
                self._workbook.get_internal_option_format(format)?,
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    /// This function writes an array formula to a cell range. In Excel an array formula is a formula that performs a calculation on a set of values.
    /// In Excel an array formula is indicated by a pair of braces around the formula: `{=SUM(A1:B1*A2:B2)}`.
    ///
    /// Array formulas can return a single value or a range or values. For array formulas that return a range of values you must specify the range that the return values will be written to. This is why this function has first_ and last_ row/column parameters:
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-worksheet_write_array_formula-1.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// worksheet.write_array_formula(4, 0, 6, 0, "{=TREND(C5:C7,B5:B7)}", None)?;
    /// # workbook.close()
    /// # }
    /// ```
    /// If the array formula returns a single value then the first_ and last_ parameters should be the same:
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-worksheet_write_array_formula-2.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// worksheet.write_array_formula(1, 0, 1, 0, "{=SUM(B1:C1*B2:C2)}", None)?;
    /// # workbook.close()
    /// # }
    /// ```
    pub fn write_array_formula(
        &mut self,
        first_row: WorksheetRow,
        first_col: WorksheetCol,
        last_row: WorksheetRow,
        last_col: WorksheetCol,
        formula: &str,
        format: Option<&Format>,
    ) -> Result<(), XlsxError> {
        let mut c_string_helper = CStringHelper::new();
        unsafe {
            let result = libxlsxwriter_sys::worksheet_write_array_formula(
                self.worksheet,
                first_row,
                first_col,
                last_row,
                last_col,
                c_string_helper.add(formula)?,
                self._workbook.get_internal_option_format(format)?,
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    /// This function can be used to write a date or time to the cell specified by row and column:
    /// ```rust
    /// use xlsxwriter::prelude::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-worksheet_write_datetime-1.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// let datetime = DateTime::new(2013, 2, 28, 12, 0, 0.0);
    /// worksheet.write_datetime(1, 0, &datetime, Some(&Format::new().set_num_format("mmm d yyyy hh:mm AM/PM")))?;
    /// # workbook.close()
    /// # }
    /// ```
    /// ![Result Image](https://github.com/informationsea/xlsxwriter-rs/raw/master/images/test-worksheet-write_datetime-1.png)
    ///
    /// The `format` parameter should be used to apply formatting to the cell using a [`Format`] object as shown above. Without a date format the datetime will appear as a number only.
    ///
    /// See [Working with Dates and Times](https://libxlsxwriter.github.io/working_with_dates.html) for more information about handling dates and times in libxlsxwriter.
    pub fn write_datetime(
        &mut self,
        row: WorksheetRow,
        col: WorksheetCol,
        datetime: &DateTime,
        format: Option<&Format>,
    ) -> Result<(), XlsxError> {
        unsafe {
            let mut xls_datetime: libxlsxwriter_sys::lxw_datetime = datetime.into();
            let result = libxlsxwriter_sys::worksheet_write_datetime(
                self.worksheet,
                row,
                col,
                &mut xls_datetime,
                self._workbook.get_internal_option_format(format)?,
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    /// This function is used to write a URL/hyperlink to a worksheet cell specified by row and column.
    /// The format parameter is used to apply formatting to the cell. This parameter can be `None` to indicate no formatting or it can be a [Format](struct.Format.html) object. The typical worksheet format for a hyperlink is a blue underline:
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-worksheet_write_url-1.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// let mut url_format = Format::new();
    /// url_format.set_underline(FormatUnderline::Single).set_font_color(FormatColor::Blue);
    /// worksheet.write_url(0, 0, "http://libxlsxwriter.github.io", Some(&url_format))?;
    /// # workbook.close()
    /// # }
    /// ```
    ///
    /// The usual web style URI's are supported: `http://`, `https://`, `ftp://` and `mailto:` :
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-worksheet_write_url-2.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # let mut url_format = Format::new();
    /// # url_format.set_underline(FormatUnderline::Single).set_font_color(FormatColor::Blue);
    /// worksheet.write_url(0, 0, "ftp://www.python.org/", Some(&url_format))?;
    /// worksheet.write_url(1, 0, "http://www.python.org/", Some(&url_format))?;
    /// worksheet.write_url(2, 0, "https://www.python.org/", Some(&url_format))?;
    /// worksheet.write_url(3, 0, "mailto:foo@example.com", Some(&url_format))?;
    /// # workbook.close()
    /// # }
    /// ```
    ///
    /// An Excel hyperlink is comprised of two elements: the displayed string and the non-displayed link. By default the displayed string is the same as the link. However, it is possible to overwrite it with any other libxlsxwriter type using the appropriate `Worksheet.write_*()` function. The most common case is to overwrite the displayed link text with another string:
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-worksheet_write_url-3.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # let mut url_format = Format::new();
    /// # url_format.set_underline(FormatUnderline::Single).set_font_color(FormatColor::Blue);
    /// worksheet.write_url(0, 0, "http://libxlsxwriter.github.io", Some(&url_format))?;
    /// worksheet.write_string(0, 0, "Read the documentation.", Some(&url_format))?;
    /// # workbook.close()
    /// # }
    /// ```
    ///
    /// Two local URIs are supported: `internal:` and `external:`. These are used for hyperlinks to internal worksheet references or external workbook and worksheet references:
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-worksheet_write_url-4.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # let mut worksheet2 = workbook.add_worksheet(None)?;
    /// # let mut worksheet3 = workbook.add_worksheet(Some("Sales Data"))?;
    /// # let mut url_format = Format::new();
    /// # url_format.set_underline(FormatUnderline::Single).set_font_color(FormatColor::Blue);
    /// worksheet.write_url(0, 0, "internal:Sheet2!A1", Some(&url_format))?;
    /// worksheet.write_url(1, 0, "internal:Sheet2!B2", Some(&url_format))?;
    /// worksheet.write_url(2, 0, "internal:Sheet2!A1:B2", Some(&url_format))?;
    /// worksheet.write_url(3, 0, "internal:'Sales Data'!A1", Some(&url_format))?;
    /// worksheet.write_url(4, 0, "external:c:\\temp\\foo.xlsx", Some(&url_format))?;
    /// worksheet.write_url(5, 0, "external:c:\\foo.xlsx#Sheet2!A1", Some(&url_format))?;
    /// worksheet.write_url(6, 0, "external:..\\foo.xlsx", Some(&url_format))?;
    /// worksheet.write_url(7, 0, "external:..\\foo.xlsx#Sheet2!A1", Some(&url_format))?;
    /// worksheet.write_url(8, 0, "external:\\\\NET\\share\\foo.xlsx", Some(&url_format))?;
    /// # workbook.close()
    /// # }
    /// ```
    pub fn write_url(
        &mut self,
        row: WorksheetRow,
        col: WorksheetCol,
        url: &str,
        format: Option<&Format>,
    ) -> Result<(), XlsxError> {
        unsafe {
            let result = libxlsxwriter_sys::worksheet_write_url(
                self.worksheet,
                row,
                col,
                CString::new(url)?.as_c_str().as_ptr(),
                self._workbook.get_internal_option_format(format)?,
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    /// Write an Excel boolean to the cell specified by row and column:
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-worksheet_write_boolean-1.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// worksheet.write_boolean(0, 0, true, None)?;
    /// worksheet.write_boolean(1, 0, false, None)?;
    /// # workbook.close()
    /// # }
    /// ```
    pub fn write_boolean(
        &mut self,
        row: WorksheetRow,
        col: WorksheetCol,
        value: bool,
        format: Option<&Format>,
    ) -> Result<(), XlsxError> {
        unsafe {
            let result = libxlsxwriter_sys::worksheet_write_boolean(
                self.worksheet,
                row,
                col,
                if value { 1 } else { 0 },
                self._workbook.get_internal_option_format(format)?,
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    /// Write a blank cell specified by row and column:
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-worksheet_write_blank-1.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// worksheet.write_blank(1, 1, None);
    /// # workbook.close()
    /// # }
    /// ```
    /// This function is used to add formatting to a cell which doesn't contain a string or number value.
    ///
    /// Excel differentiates between an "Empty" cell and a "Blank" cell. An Empty cell is a cell which doesn't contain data or formatting whilst a Blank cell doesn't contain data but does contain formatting. Excel stores Blank cells but ignores Empty cells.
    ///
    /// As such, if you write an empty cell without formatting it is ignored.
    pub fn write_blank(
        &mut self,
        row: WorksheetRow,
        col: WorksheetCol,
        format: Option<&Format>,
    ) -> Result<(), XlsxError> {
        unsafe {
            let result = libxlsxwriter_sys::worksheet_write_blank(
                self.worksheet,
                row,
                col,
                self._workbook.get_internal_option_format(format)?,
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    /// This function writes a formula or Excel function to the cell specified by row and column with a user defined numeric result:
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-worksheet_write_formula_num-1.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// worksheet.write_formula_num(1, 1, "=1 + 2", None, 3.0);
    /// # workbook.close()
    /// # }
    /// ```
    /// Libxlsxwriter doesn't calculate the value of a formula and instead stores the value 0 as the formula result.
    /// It then sets a global flag in the XLSX file to say that all formulas and functions should be recalculated when the file is opened.
    ///
    /// This is the method recommended in the Excel documentation and in general it works fine with spreadsheet applications.
    ///
    /// However, applications that don't have a facility to calculate formulas, such as Excel Viewer, or some mobile
    /// applications will only display the 0 results.
    ///
    /// If required, the worksheet_write_formula_num() function can be used to specify a formula and its result.
    ///
    /// This function is rarely required and is only provided for compatibility with some third party applications.
    /// For most applications the worksheet_write_formula() function is the recommended way of writing formulas.
    #[allow(clippy::too_many_arguments)]
    pub fn write_formula_num(
        &mut self,
        row: WorksheetRow,
        col: WorksheetCol,
        formula: &str,
        format: Option<&Format>,
        number: f64,
    ) -> Result<(), XlsxError> {
        unsafe {
            let result = libxlsxwriter_sys::worksheet_write_formula_num(
                self.worksheet,
                row,
                col,
                CString::new(formula)?.as_c_str().as_ptr(),
                self._workbook.get_internal_option_format(format)?,
                number,
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    /// This function writes a formula or Excel function to the cell specified by row and column with a user defined string result:
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-worksheet_write_formula_str-1.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// worksheet.write_formula_str(1, 1, "=\"A\" & \"B\"", None, "AB");
    /// # workbook.close()
    /// # }
    /// ```
    /// The worksheet_write_formula_str() function is similar to the worksheet_write_formula_num() function except it
    /// writes a string result instead or a numeric result. See worksheet_write_formula_num() for more details on
    /// why/when these functions are required.
    ///
    /// One place where the worksheet_write_formula_str() function may be required is to specify an empty result which
    /// will force a recalculation of the formula when loaded in LibreOffice.
    #[allow(clippy::too_many_arguments)]
    pub fn write_formula_str(
        &mut self,
        row: WorksheetRow,
        col: WorksheetCol,
        formula: &str,
        format: Option<&Format>,
        result: &str,
    ) -> Result<(), XlsxError> {
        unsafe {
            let result = libxlsxwriter_sys::worksheet_write_formula_str(
                self.worksheet,
                row,
                col,
                CString::new(formula)?.as_c_str().as_ptr(),
                self._workbook.get_internal_option_format(format)?,
                CString::new(result).unwrap().as_c_str().as_ptr(),
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    /// This function is used to write strings with multiple formats. For example to write the string 'This is bold and this is italic' you would use the following:
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-worksheet_write_richtext-1.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// worksheet.write_rich_string(
    ///     0, 0,
    ///     &[
    ///         ("This is ", None),
    ///         ("bold", Some(&Format::new().set_bold())),
    ///         (" and this is ", None),
    ///         ("italic", Some(&Format::new().set_italic()))
    ///     ],
    ///     None
    /// )?;
    /// # workbook.close()
    /// # }
    /// ```
    /// ![Result Image](https://github.com/informationsea/xlsxwriter-rs/raw/master/images/test-worksheet-write_richtext-1.png)
    ///
    /// The basic rule is to break the string into fragments and put a lxw_format object before the fragment that you want to format. So if we look at the above example again:
    ///
    /// This is **bold** and this is *italic*
    ///
    /// The would be broken down into 4 fragments:
    /// ```text
    /// default: |This is |
    /// bold:    |bold|
    /// default: | and this is |
    /// italic:  |italic|
    /// ```
    /// This in then converted to the tuple fragments shown in the example above. For the default format we use None.
    ///
    /// ### Note
    ///  Excel doesn't allow the use of two consecutive formats in a rich string or an empty string fragment. For either of these conditions a warning is raised and the input to `worksheet.write_rich_string()` is ignored.
    pub fn write_rich_string(
        &mut self,
        row: WorksheetRow,
        col: WorksheetCol,
        text: &[(&str, Option<&Format>)],
        format: Option<&Format>,
    ) -> Result<(), XlsxError> {
        let mut c_str: Vec<Vec<u8>> = text
            .iter()
            .map(|x| {
                CString::new(x.0)
                    .unwrap()
                    .as_c_str()
                    .to_bytes_with_nul()
                    .to_vec()
            })
            .collect();

        let mut rich_text: Vec<_> = text
            .iter()
            .zip(c_str.iter_mut())
            .map(|(x, y)| libxlsxwriter_sys::lxw_rich_string_tuple {
                format: self._workbook.get_internal_option_format(x.1).unwrap(), // Fix here
                string: y.as_mut_ptr() as *mut c_char,
            })
            .collect();
        let mut rich_text_ptr: Vec<*mut libxlsxwriter_sys::lxw_rich_string_tuple> = rich_text
            .iter_mut()
            .map(|x| x as *mut libxlsxwriter_sys::lxw_rich_string_tuple)
            .collect();
        rich_text_ptr.push(std::ptr::null_mut());

        unsafe {
            let result = libxlsxwriter_sys::worksheet_write_rich_string(
                self.worksheet,
                row,
                col,
                rich_text_ptr.as_mut_ptr(),
                self._workbook.get_internal_option_format(format)?,
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    pub fn set_row(
        &mut self,
        row: WorksheetRow,
        height: f64,
        format: Option<&Format>,
    ) -> Result<(), XlsxError> {
        unsafe {
            let result = libxlsxwriter_sys::worksheet_set_row(
                self.worksheet,
                row,
                height,
                self._workbook.get_internal_option_format(format)?,
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    pub fn set_row_opt(
        &mut self,
        row: WorksheetRow,
        height: f64,
        format: Option<&Format>,
        options: &RowColOptions,
    ) -> Result<(), XlsxError> {
        unsafe {
            let mut options = options.into_internal();
            let result = libxlsxwriter_sys::worksheet_set_row_opt(
                self.worksheet,
                row,
                height,
                self._workbook.get_internal_option_format(format)?,
                &mut options,
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    /// The set_row_pixels() function is the same as the [Worksheet::set_row()] function except that the height can be set in pixels.
    ///
    /// If you wish to set the format of a row without changing the height you can pass the default row height in pixels: [LXW_DEF_ROW_HEIGHT_PIXELS].
    pub fn set_row_pixels(
        &mut self,
        row: WorksheetRow,
        pixels: u32,
        format: Option<&Format>,
    ) -> Result<(), XlsxError> {
        unsafe {
            let result = libxlsxwriter_sys::worksheet_set_row_pixels(
                self.worksheet,
                row,
                pixels,
                self._workbook.get_internal_option_format(format)?,
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    pub fn set_row_pixels_opt(
        &mut self,
        row: WorksheetRow,
        pixels: u32,
        format: Option<&Format>,
        options: &RowColOptions,
    ) -> Result<(), XlsxError> {
        let mut options = options.into_internal();
        unsafe {
            let result = libxlsxwriter_sys::worksheet_set_row_pixels_opt(
                self.worksheet,
                row,
                pixels,
                self._workbook.get_internal_option_format(format)?,
                &mut options,
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    pub fn set_column(
        &mut self,
        first_col: WorksheetCol,
        last_col: WorksheetCol,
        width: f64,
        format: Option<&Format>,
    ) -> Result<(), XlsxError> {
        unsafe {
            let result = libxlsxwriter_sys::worksheet_set_column(
                self.worksheet,
                first_col,
                last_col,
                width,
                self._workbook.get_internal_option_format(format)?,
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    pub fn set_column_opt(
        &mut self,
        first_col: WorksheetCol,
        last_col: WorksheetCol,
        width: f64,
        format: Option<&Format>,
        options: &RowColOptions,
    ) -> Result<(), XlsxError> {
        let mut options = options.into_internal();
        unsafe {
            let result = libxlsxwriter_sys::worksheet_set_column_opt(
                self.worksheet,
                first_col,
                last_col,
                width,
                self._workbook.get_internal_option_format(format)?,
                &mut options,
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    pub fn set_column_pixels(
        &mut self,
        first_col: WorksheetCol,
        last_col: WorksheetCol,
        pixels: u32,
        format: Option<&Format>,
    ) -> Result<(), XlsxError> {
        unsafe {
            let result = libxlsxwriter_sys::worksheet_set_column_pixels(
                self.worksheet,
                first_col,
                last_col,
                pixels,
                self._workbook.get_internal_option_format(format)?,
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    pub fn set_column_pixels_opt(
        &mut self,
        first_col: WorksheetCol,
        last_col: WorksheetCol,
        pixels: u32,
        format: Option<&Format>,
        options: &mut RowColOptions,
    ) -> Result<(), XlsxError> {
        let mut options = options.into_internal();
        unsafe {
            let result = libxlsxwriter_sys::worksheet_set_column_pixels_opt(
                self.worksheet,
                first_col,
                last_col,
                pixels,
                self._workbook.get_internal_option_format(format)?,
                &mut options,
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    /// This function can be used to insert a image into a worksheet. The image can be in PNG, JPEG or BMP format:
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-worksheet_insert_image-1.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// worksheet.insert_image(2, 1, "../images/simple1.png")?;
    /// # workbook.close()
    /// # }
    /// ```
    /// ![Result Image](https://github.com/informationsea/xlsxwriter-rs/raw/master/images/test-worksheet-insert_image-1.png)
    ///
    /// The Worksheet.insert_image_opt() function takes additional optional parameters to position and scale the image, see below.
    ///
    /// ### Note
    /// The scaling of a image may be affected if is crosses a row that has its default height changed due to a font that is larger than
    /// the default font size or that has text wrapping turned on. To avoid this you should explicitly set the height of the row using
    /// [`Worksheet::set_row`] if it crosses an inserted image.
    ///
    /// BMP images are only supported for backward compatibility. In general it is best to avoid BMP images since they aren't compressed.
    /// If used, BMP images must be 24 bit, true color, bitmaps.
    pub fn insert_image(
        &mut self,
        row: WorksheetRow,
        col: WorksheetCol,
        filename: &str,
    ) -> Result<(), XlsxError> {
        unsafe {
            let result = libxlsxwriter_sys::worksheet_insert_image(
                self.worksheet,
                row,
                col,
                CString::new(filename).unwrap().as_c_str().as_ptr(),
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    /// This function is like Worksheet.insert_image() function except that it takes an optional `ImageOptions` struct to scale and position the image:
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-worksheet_insert_image_opt-1.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// worksheet.insert_image_opt(
    ///     2, 1,
    ///    "../images/simple1.png",
    ///     &ImageOptions{
    ///         x_offset: 30,
    ///         y_offset: 30,
    ///         x_scale: 0.5,
    ///         y_scale: 0.5,
    ///     }
    /// )?;
    /// # workbook.close()
    /// # }
    /// ```
    /// ![Result Image](https://github.com/informationsea/xlsxwriter-rs/raw/master/images/test-worksheet-insert_image_opt-1.png)
    ///
    /// ### Note
    /// See the notes about row scaling and BMP images in [`Worksheet::insert_image`] above.
    pub fn insert_image_opt(
        &mut self,
        row: WorksheetRow,
        col: WorksheetCol,
        filename: &str,
        opt: &ImageOptions,
    ) -> Result<(), XlsxError> {
        let mut opt_struct = opt.into();
        unsafe {
            let result = libxlsxwriter_sys::worksheet_insert_image_opt(
                self.worksheet,
                row,
                col,
                CString::new(filename).unwrap().as_c_str().as_ptr(),
                &mut opt_struct,
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    /// This function can be used to insert a image into a worksheet from a memory buffer:
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-worksheet_insert_image_buffer-1.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// let data = include_bytes!("../../../images/simple1.png");
    /// worksheet.insert_image_buffer(0, 0, &data[..])?;
    /// # workbook.close()
    /// # }
    /// ```
    /// See [`Worksheet::insert_image`] for details about the supported image formats, and other image features.
    pub fn insert_image_buffer(
        &mut self,
        row: WorksheetRow,
        col: WorksheetCol,
        buffer: &[u8],
    ) -> Result<(), XlsxError> {
        unsafe {
            let result = libxlsxwriter_sys::worksheet_insert_image_buffer(
                self.worksheet,
                row,
                col,
                buffer.as_ptr(),
                buffer.len(),
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    pub fn insert_image_buffer_opt(
        &mut self,
        row: WorksheetRow,
        col: WorksheetCol,
        buffer: &[u8],
        opt: &ImageOptions,
    ) -> Result<(), XlsxError> {
        let mut opt_struct = opt.into();
        unsafe {
            let result = libxlsxwriter_sys::worksheet_insert_image_buffer_opt(
                self.worksheet,
                row,
                col,
                buffer.as_ptr(),
                buffer.len(),
                &mut opt_struct,
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    pub fn insert_chart(
        &mut self,
        row: WorksheetRow,
        column: WorksheetCol,
        chart: &Chart,
    ) -> Result<(), XlsxError> {
        unsafe {
            let result =
                libxlsxwriter_sys::worksheet_insert_chart(self.worksheet, row, column, chart.chart);
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    /// The [`Worksheet::merge_range`] function allows cells to be merged together so that they act as a single area.
    pub fn merge_range(
        &mut self,
        first_row: WorksheetRow,
        first_col: WorksheetCol,
        last_row: WorksheetRow,
        last_col: WorksheetCol,
        string: &str,
        format: Option<&Format>,
    ) -> Result<(), XlsxError> {
        unsafe {
            let result = libxlsxwriter_sys::worksheet_merge_range(
                self.worksheet,
                first_row,
                first_col,
                last_row,
                last_col,
                CString::new(string)?.as_c_str().as_ptr(),
                self._workbook.get_internal_option_format(format)?,
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    /// The [`Worksheet::activate`] function is used to specify which worksheet is initially visible in a multi-sheet workbook.
    pub fn activate(&mut self) {
        unsafe {
            libxlsxwriter_sys::worksheet_activate(self.worksheet);
        }
    }

    /// The [`Worksheet::select`] function is used to indicate that a worksheet is selected in a multi-sheet workbook.
    ///
    /// A selected worksheet has its tab highlighted. Selecting worksheets is a way of grouping them together so that, for example, several worksheets could be printed in one go.
    /// A worksheet that has been activated via the [`Worksheet::activate`] function will also appear as selected.
    pub fn select(&mut self) {
        unsafe {
            libxlsxwriter_sys::worksheet_select(self.worksheet);
        }
    }

    /// The [`Worksheet::hide`] function is used to hide a worksheet.
    ///
    /// You may wish to hide a worksheet in order to avoid confusing a user with intermediate data or calculations.
    ///
    /// A hidden worksheet can not be activated or selected so this function is mutually exclusive with the [`Worksheet::activate`]
    /// and [`Worksheet::select`] functions. In addition, since the first worksheet will default to being the active worksheet,
    /// you cannot hide the first worksheet without activating another sheet.
    pub fn hide(&mut self) {
        unsafe {
            libxlsxwriter_sys::worksheet_hide(self.worksheet);
        }
    }

    /// The [`Worksheet::activate`] function determines which worksheet is initially selected. However,
    /// if there are a large number of worksheets the selected worksheet may not appear on the screen.
    /// To avoid this you can select the leftmost visible worksheet tab using [`Worksheet::set_first_sheet`]
    ///
    /// This function is not required very often. The default value is the first worksheet.
    pub fn set_first_sheet(&mut self) {
        unsafe {
            libxlsxwriter_sys::worksheet_set_first_sheet(self.worksheet);
        }
    }

    /// The [`Worksheet::freeze_panes`] function can be used to divide a worksheet into horizontal or
    /// vertical regions known as panes and to "freeze" these panes so that the splitter bars are not visible.
    ///
    /// The parameters row and col are used to specify the location of the split. It should be noted that the
    /// split is specified at the top or left of a cell and that the function uses zero based indexing. Therefore
    /// to freeze the first row of a worksheet it is necessary to specify the split at row 2 (which is 1 as the zero-based index).
    ///
    /// You can set one of the row and col parameters as zero if you do not want either a vertical or horizontal split.
    pub fn freeze_panes(&mut self, row: WorksheetRow, col: WorksheetCol) {
        unsafe {
            libxlsxwriter_sys::worksheet_freeze_panes(self.worksheet, row, col);
        }
    }

    /// The [`Worksheet::split_panes`] function can be used to divide a worksheet into horizontal or vertical regions known as panes.
    /// This function is different from the [`Worksheet::freeze_panes`] function in that the splits between the panes will be visible
    /// to the user and each pane will have its own scroll bars.
    ///
    /// The parameters vertical and horizontal are used to specify the vertical and horizontal position of the split. The units for
    /// vertical and horizontal are the same as those used by Excel to specify row height and column width. However, the vertical
    /// and horizontal units are different from each other. Therefore you must specify the vertical and horizontal parameters in
    /// terms of the row heights and column widths that you have set or the default values which are 15 for a row and 8.43 for a column.
    pub fn split_panes(&mut self, vertical: f64, horizontal: f64) {
        unsafe {
            libxlsxwriter_sys::worksheet_split_panes(self.worksheet, vertical, horizontal);
        }
    }

    /// The [`Worksheet::set_selection`] function can be used to specify which cell or range of cells is selected in a worksheet:
    /// The most common requirement is to select a single cell, in which case the first_ and last_ parameters should be the same.
    ///
    /// The active cell within a selected range is determined by the order in which `first_` and `last_` are specified.
    pub fn set_selection(
        &mut self,
        first_row: WorksheetRow,
        first_col: WorksheetCol,
        last_row: WorksheetRow,
        last_col: WorksheetCol,
    ) {
        unsafe {
            libxlsxwriter_sys::worksheet_set_selection(
                self.worksheet,
                first_row,
                first_col,
                last_row,
                last_col,
            );
        }
    }

    pub fn set_landscape(&mut self) {
        unsafe {
            libxlsxwriter_sys::worksheet_set_landscape(self.worksheet);
        }
    }

    pub fn set_portrait(&mut self) {
        unsafe {
            libxlsxwriter_sys::worksheet_set_portrait(self.worksheet);
        }
    }

    pub fn set_page_view(&mut self) {
        unsafe {
            libxlsxwriter_sys::worksheet_set_page_view(self.worksheet);
        }
    }

    pub fn set_paper(&mut self, paper: PaperType) {
        unsafe {
            libxlsxwriter_sys::worksheet_set_paper(self.worksheet, paper.value());
        }
    }

    pub fn set_header(&mut self, header: &str) -> Result<(), XlsxError> {
        unsafe {
            let result = libxlsxwriter_sys::worksheet_set_header(
                self.worksheet,
                CString::new(header).unwrap().as_c_str().as_ptr(),
            );

            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    pub fn set_footer(&mut self, footer: &str) -> Result<(), XlsxError> {
        unsafe {
            let result = libxlsxwriter_sys::worksheet_set_footer(
                self.worksheet,
                CString::new(footer).unwrap().as_c_str().as_ptr(),
            );

            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    pub fn set_header_opt(
        &mut self,
        header: &str,
        options: &HeaderFooterOptions,
    ) -> Result<(), XlsxError> {
        unsafe {
            let result = libxlsxwriter_sys::worksheet_set_header_opt(
                self.worksheet,
                CString::new(header).unwrap().as_c_str().as_ptr(),
                &mut options.into(),
            );

            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    pub fn set_footer_opt(
        &mut self,
        footer: &str,
        options: &HeaderFooterOptions,
    ) -> Result<(), XlsxError> {
        unsafe {
            let result = libxlsxwriter_sys::worksheet_set_footer_opt(
                self.worksheet,
                CString::new(footer).unwrap().as_c_str().as_ptr(),
                &mut options.into(),
            );

            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    pub fn set_h_pagebreaks(&mut self, breaks: &[WorksheetRow]) -> Result<(), XlsxError> {
        let mut breaks_vec = breaks.to_vec();
        breaks_vec.push(0);
        unsafe {
            let result = libxlsxwriter_sys::worksheet_set_h_pagebreaks(
                self.worksheet,
                breaks_vec.as_mut_ptr(),
            );

            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    pub fn set_v_pagebreaks(&mut self, breaks: &[WorksheetCol]) -> Result<(), XlsxError> {
        let mut breaks_vec = breaks.to_vec();
        breaks_vec.push(0);
        unsafe {
            let result = libxlsxwriter_sys::worksheet_set_v_pagebreaks(
                self.worksheet,
                breaks_vec.as_mut_ptr(),
            );

            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    pub fn print_across(&mut self) {
        unsafe {
            libxlsxwriter_sys::worksheet_print_across(self.worksheet);
        }
    }

    pub fn set_zoom(&mut self, scale: u16) {
        unsafe {
            libxlsxwriter_sys::worksheet_set_zoom(self.worksheet, scale);
        }
    }

    pub fn gridlines(&mut self, option: GridLines) {
        unsafe {
            libxlsxwriter_sys::worksheet_gridlines(self.worksheet, option.value());
        }
    }

    pub fn center_horizontally(&mut self) {
        unsafe {
            libxlsxwriter_sys::worksheet_center_horizontally(self.worksheet);
        }
    }

    pub fn center_vertically(&mut self) {
        unsafe {
            libxlsxwriter_sys::worksheet_center_vertically(self.worksheet);
        }
    }

    pub fn print_row_col_headers(&mut self) {
        unsafe {
            libxlsxwriter_sys::worksheet_print_row_col_headers(self.worksheet);
        }
    }

    pub fn repeat_rows(
        &mut self,
        first_row: WorksheetRow,
        last_row: WorksheetRow,
    ) -> Result<(), XlsxError> {
        unsafe {
            let result =
                libxlsxwriter_sys::worksheet_repeat_rows(self.worksheet, first_row, last_row);
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    pub fn repeat_columns(
        &mut self,
        first_col: WorksheetCol,
        last_col: WorksheetCol,
    ) -> Result<(), XlsxError> {
        unsafe {
            let result =
                libxlsxwriter_sys::worksheet_repeat_columns(self.worksheet, first_col, last_col);
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    pub fn print_area(
        &mut self,
        first_row: WorksheetRow,
        first_col: WorksheetCol,
        last_row: WorksheetRow,
        last_col: WorksheetCol,
    ) -> Result<(), XlsxError> {
        unsafe {
            let result = libxlsxwriter_sys::worksheet_print_area(
                self.worksheet,
                first_row,
                first_col,
                last_row,
                last_col,
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    pub fn fit_to_pages(&mut self, width: u16, height: u16) {
        unsafe {
            libxlsxwriter_sys::worksheet_fit_to_pages(self.worksheet, width, height);
        }
    }

    pub fn set_start_page(&mut self, start_page: u16) {
        unsafe {
            libxlsxwriter_sys::worksheet_set_start_page(self.worksheet, start_page);
        }
    }

    pub fn set_print_scale(&mut self, scale: u16) {
        unsafe {
            libxlsxwriter_sys::worksheet_set_print_scale(self.worksheet, scale);
        }
    }

    pub fn set_right_to_left(&mut self) {
        unsafe {
            libxlsxwriter_sys::worksheet_right_to_left(self.worksheet);
        }
    }

    pub fn set_hide_zero(&mut self) {
        unsafe {
            libxlsxwriter_sys::worksheet_hide_zero(self.worksheet);
        }
    }

    pub fn set_tab_color(&mut self, color: FormatColor) {
        unsafe {
            libxlsxwriter_sys::worksheet_set_tab_color(self.worksheet, color.value());
        }
    }

    pub fn protect(&mut self, password: &str, protection: &Protection) {
        unsafe {
            libxlsxwriter_sys::worksheet_protect(
                self.worksheet,
                CString::new(password).unwrap().as_c_str().as_ptr(),
                &mut protection.into(),
            );
        }
    }

    pub fn outline_settings(
        &mut self,
        visible: bool,
        symbols_below: bool,
        symbols_right: bool,
        auto_style: bool,
    ) {
        unsafe {
            libxlsxwriter_sys::worksheet_outline_settings(
                self.worksheet,
                convert_bool(visible),
                convert_bool(symbols_below),
                convert_bool(symbols_right),
                convert_bool(auto_style),
            )
        }
    }

    pub fn set_default_row(&mut self, height: f64, hide_unused_rows: bool) {
        unsafe {
            libxlsxwriter_sys::worksheet_set_default_row(
                self.worksheet,
                height,
                convert_bool(hide_unused_rows),
            )
        }
    }

    pub fn set_vba_name(&mut self, name: &str) -> Result<(), XlsxError> {
        unsafe {
            let result = libxlsxwriter_sys::worksheet_set_vba_name(
                self.worksheet,
                CString::new(name).unwrap().as_c_str().as_ptr(),
            );

            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }
}
