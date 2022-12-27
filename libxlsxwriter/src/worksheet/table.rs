use std::os::raw::c_char;

use crate::{
    convert_bool, error::XlsxErrorSource, try_to_vec, CStringHelper, WorksheetCol, WorksheetRow,
    XlsxError,
};

use super::{Format, Worksheet};

/// Structure to set the options of a table column.
///
/// Please read [libxslxwriter document](https://libxlsxwriter.github.io/working_with_tables.html) to learn more.
#[derive(Default)]
pub struct TableColumn {
    /// Set the header name/caption for the column. If NULL the header defaults to Column 1, Column 2, etc.
    pub header: Option<String>,

    /// Set the formula for the column.
    pub formula: Option<String>,

    /// Set the string description for the column total.
    pub total_string: Option<String>,

    /// Set the function for the column total.
    pub total_function: TableTotalFunction,

    /// Set the format for the column header.
    pub header_format: Option<Format>,

    /// Set the format for the data rows in the column.
    pub format: Option<Format>,

    /// Set the formula value for the column total (not generally required).
    pub total_value: f64,
}

/// The type of table style (Light, Medium or Dark).
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub enum TableStyleType {
    Default,
    Light,
    Medium,
    Dark,
}

impl Default for TableStyleType {
    fn default() -> TableStyleType {
        TableStyleType::Default
    }
}

impl From<TableStyleType> for u8 {
    fn from(t: TableStyleType) -> u8 {
        (match t {
            TableStyleType::Dark => {
                libxlsxwriter_sys::lxw_table_style_type_LXW_TABLE_STYLE_TYPE_DARK
            }
            TableStyleType::Light => {
                libxlsxwriter_sys::lxw_table_style_type_LXW_TABLE_STYLE_TYPE_LIGHT
            }
            TableStyleType::Medium => {
                libxlsxwriter_sys::lxw_table_style_type_LXW_TABLE_STYLE_TYPE_MEDIUM
            }
            TableStyleType::Default => {
                libxlsxwriter_sys::lxw_table_style_type_LXW_TABLE_STYLE_TYPE_DEFAULT
            }
        }) as u8
    }
}

/// Definitions for the standard Excel functions that are available via the dropdown in the total row of an Excel table.
///
/// Please read [libxslxwriter document](https://libxlsxwriter.github.io/working_with_tables.html) to learn more.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub enum TableTotalFunction {
    None,

    /// Use the average function as the table total.
    Average,

    /// Use the count numbers function as the table total.
    CountNums,

    /// Use the count function as the table total.
    Count,

    /// Use the max function as the table total.
    Max,

    /// Use the min function as the table total.
    Min,

    /// Use the standard deviation function as the table total.
    StdDev,

    /// Use the sum function as the table total.
    Sum,

    /// Use the var function as the table total.
    Var,
}

impl Default for TableTotalFunction {
    fn default() -> TableTotalFunction {
        TableTotalFunction::None
    }
}

impl From<TableTotalFunction> for u8 {
    fn from(f: TableTotalFunction) -> u8 {
        (match f {
            TableTotalFunction::None => 0,
            TableTotalFunction::Average => {
                libxlsxwriter_sys::lxw_table_total_functions_LXW_TABLE_FUNCTION_AVERAGE
            }
            TableTotalFunction::CountNums => {
                libxlsxwriter_sys::lxw_table_total_functions_LXW_TABLE_FUNCTION_COUNT_NUMS
            }
            TableTotalFunction::Count => {
                libxlsxwriter_sys::lxw_table_total_functions_LXW_TABLE_FUNCTION_COUNT
            }
            TableTotalFunction::Max => {
                libxlsxwriter_sys::lxw_table_total_functions_LXW_TABLE_FUNCTION_MAX
            }
            TableTotalFunction::Min => {
                libxlsxwriter_sys::lxw_table_total_functions_LXW_TABLE_FUNCTION_MIN
            }
            TableTotalFunction::StdDev => {
                libxlsxwriter_sys::lxw_table_total_functions_LXW_TABLE_FUNCTION_STD_DEV
            }
            TableTotalFunction::Sum => {
                libxlsxwriter_sys::lxw_table_total_functions_LXW_TABLE_FUNCTION_SUM
            }
            TableTotalFunction::Var => {
                libxlsxwriter_sys::lxw_table_total_functions_LXW_TABLE_FUNCTION_VAR
            }
        }) as u8
    }
}

/// Options used to define worksheet tables.
/// ```rust
/// # use xlsxwriter::prelude::*;
/// # fn main() -> Result<(), XlsxError> {
/// # let workbook = Workbook::new("test-worksheet_add_table-1.xlsx")?;
/// # let mut worksheet = workbook.add_worksheet(None)?;
/// worksheet.write_string(0, 0, "header 1", None)?;
/// worksheet.write_string(0, 1, "header 2", None)?;
/// worksheet.write_string(1, 0, "content 1", None)?;
/// worksheet.write_number(1, 1, 1.0, None)?;
/// worksheet.write_string(2, 0, "content 2", None)?;
/// worksheet.write_number(2, 1, 2.0, None)?;
/// worksheet.write_string(3, 0, "content 3", None)?;
/// worksheet.write_number(3, 1, 3.0, None)?;
/// worksheet.add_table(0, 0, 3, 1, None)?;
/// # workbook.close()
/// # }
/// ```
#[derive(Default)]
pub struct TableOptions {
    /**
     * The name parameter is used to set the name of the table. This parameter is optional and by
     * default tables are named Table1, Table2, etc. in the worksheet order that they are added.
     * If you override the table name you must ensure that it doesn't clash with an existing table
     * name and that it follows Excel's requirements for table names, see the Microsoft Office documentation.
     */
    pub name: Option<String>,

    /// The `no_header_row` parameter can be used to turn off the header row in the table. It is on by default.    
    pub no_header_row: bool,

    /// The `no_autofilter` parameter can be used to turn off the autofilter in the header row. It is on by default.    
    pub no_autofilter: bool,

    /// The `no_banded_rows` parameter can be used to turn off the rows of alternating color in the table. It is on by default.
    pub no_banded_rows: bool,

    /// The `banded_columns` parameter can be used to used to create columns of alternating color in the table. It is off by default.
    pub banded_columns: bool,

    /// The `first_column` parameter can be used to highlight the first column of the table. The type of highlighting will depend on the style_type of the table. It may be bold text or a different color. It is off by default.
    pub first_column: bool,

    /// The `last_column` parameter can be used to highlight the last column of the table. The type of highlighting will depend on the style of the table. It may be bold text or a different color. It is off by default.
    pub last_column: bool,

    /// The `style_type` parameter can be used to set the style of the table, in conjunction with the style_type_number parameter.
    pub style_type: TableStyleType,

    /// The `style_type_number` parameter is used with style_type to set the style of a worksheet table.     
    pub style_type_number: u8,

    /// The `total_row` parameter can be used to turn on the total row in the last row of a table. It is distinguished from the other rows by a different formatting and also with dropdown SUBTOTAL functions.
    pub total_row: bool,

    /// The columns parameter can be used to set properties for columns within the table.
    pub columns: Option<Vec<TableColumn>>,
}

impl<'a> Worksheet<'a> {
    /// This function is used to add a table to a worksheet. Tables in Excel are a way of grouping a
    /// range of cells into a single entity that has common formatting or that can be referenced
    /// from formulas. Tables can have column headers, autofilters, total rows, column formulas and
    /// default formatting.
    ///
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-worksheet_add_table-0.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// worksheet.write_string(0, 0, "header 1", None)?;
    /// worksheet.write_string(0, 1, "header 2", None)?;
    /// worksheet.write_string(1, 0, "content 1", None)?;
    /// worksheet.write_number(1, 1, 1.0, None)?;
    /// worksheet.write_string(2, 0, "content 2", None)?;
    /// worksheet.write_number(2, 1, 2.0, None)?;
    /// worksheet.write_string(3, 0, "content 3", None)?;
    /// worksheet.write_number(3, 1, 3.0, None)?;
    /// worksheet.add_table(0, 0, 3, 1, None)?;
    /// # workbook.close()
    /// # }
    /// ```
    ///
    /// Please read [libxslxwriter document](https://libxlsxwriter.github.io/working_with_tables.html) to learn more.
    pub fn add_table(
        &mut self,
        first_row: WorksheetRow,
        first_col: WorksheetCol,
        last_row: WorksheetRow,
        last_col: WorksheetCol,
        options: Option<TableOptions>,
    ) -> Result<(), XlsxError> {
        let mut cstring_helper = CStringHelper::new();

        if options
            .as_ref()
            .map(|x| x.columns.as_ref())
            .flatten()
            .map(|x| x.len() as WorksheetCol != last_col - first_col + 1)
            .unwrap_or(false)
        {
            return Err(XlsxError {
                source: XlsxErrorSource::NumberOfColumnsIsNotMatched,
            });
        }

        let columns: Option<Vec<_>> = options
            .as_ref()
            .map(|x| x.columns.as_ref())
            .flatten()
            .map(|x| {
                try_to_vec(
                    x.iter()
                        .map(
                            |y| -> Result<libxlsxwriter_sys::lxw_table_column, XlsxError> {
                                Ok(libxlsxwriter_sys::lxw_table_column {
                                    header: cstring_helper.add_opt(y.header.as_deref())?
                                        as *mut c_char,
                                    formula: cstring_helper.add_opt(y.formula.as_deref())?
                                        as *mut c_char,
                                    total_string: cstring_helper
                                        .add_opt(y.total_string.as_deref())?
                                        as *mut c_char,
                                    total_function: y.total_function.into(),
                                    header_format: self
                                        ._workbook
                                        .get_internal_option_format(y.header_format.as_ref())
                                        .unwrap(), // fix here
                                    format: self
                                        ._workbook
                                        .get_internal_option_format(y.format.as_ref())
                                        .unwrap(), // fix here
                                    total_value: y.total_value,
                                })
                            },
                        )
                        .map(|x| x.map(|y| Box::pin(y))),
                )
            })
            .transpose()?;

        let columns_ptr: Option<Vec<_>> = columns.as_ref().map(|x| {
            let mut p: Vec<_> = x
                .iter()
                .map(|x| x.as_ref().get_ref() as *const libxlsxwriter_sys::lxw_table_column)
                .map(|x| x as *mut libxlsxwriter_sys::lxw_table_column)
                .collect();
            p.push(std::ptr::null_mut());
            p
        });

        let mut options = if let Some(options) = options {
            Some(libxlsxwriter_sys::lxw_table_options {
                name: cstring_helper.add_opt(options.name.as_deref())? as *mut c_char,
                no_header_row: convert_bool(options.no_header_row),
                no_autofilter: convert_bool(options.no_autofilter),
                no_banded_rows: convert_bool(options.no_banded_rows),
                banded_columns: convert_bool(options.banded_columns),
                first_column: convert_bool(options.first_column),
                last_column: convert_bool(options.last_column),
                style_type: options.style_type.into(),
                style_type_number: options.style_type_number,
                total_row: convert_bool(options.total_row),
                columns: columns_ptr
                    .as_ref()
                    .map(|x| x.as_ptr() as *mut *mut libxlsxwriter_sys::lxw_table_column)
                    .unwrap_or(std::ptr::null_mut()),
            })
        } else {
            None
        };

        unsafe {
            let result = libxlsxwriter_sys::worksheet_add_table(
                self.worksheet,
                first_row,
                first_col,
                last_row,
                last_col,
                options
                    .as_mut()
                    .map(|x| x as *mut libxlsxwriter_sys::lxw_table_options)
                    .unwrap_or(std::ptr::null_mut()),
            );
            std::mem::drop(options);
            std::mem::drop(columns_ptr);
            std::mem::drop(columns);

            std::mem::drop(cstring_helper);

            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }
}
