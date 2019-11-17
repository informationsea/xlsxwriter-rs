use super::{convert_bool, DataValidation, Format, FormatColor, Workbook, XlsxError};
use std::ffi::CString;

#[derive(Debug, Clone, PartialEq, PartialOrd)]
pub struct DateTime {
    pub year: i16,
    pub month: i8,
    pub day: i8,
    pub hour: i8,
    pub min: i8,
    pub second: f64,
}

impl DateTime {
    pub fn new(year: i16, month: i8, day: i8, hour: i8, min: i8, second: f64) -> DateTime {
        DateTime {
            year,
            month,
            day,
            hour,
            min,
            second,
        }
    }
}

impl Into<libxlsxwriter_sys::lxw_datetime> for &DateTime {
    fn into(self) -> libxlsxwriter_sys::lxw_datetime {
        libxlsxwriter_sys::lxw_datetime {
            year: self.year.into(),
            month: self.month.into(),
            day: self.day.into(),
            hour: self.hour.into(),
            min: self.min.into(),
            sec: self.second,
        }
    }
}

#[derive(Debug, Clone, PartialEq, PartialOrd)]
pub struct ImageOptions {
    pub x_offset: i32,
    pub y_offset: i32,
    pub x_scale: f64,
    pub y_scale: f64,
}

impl Into<libxlsxwriter_sys::lxw_image_options> for &ImageOptions {
    fn into(self) -> libxlsxwriter_sys::lxw_image_options {
        libxlsxwriter_sys::lxw_image_options {
            x_offset: self.x_offset,
            y_offset: self.y_offset,
            x_scale: self.x_scale,
            y_scale: self.y_scale,
            row: 0,
            col: 0,
            filename: std::ptr::null_mut(),
            description: std::ptr::null_mut(),
            url: std::ptr::null_mut(),
            tip: std::ptr::null_mut(),
            anchor: 0,
            stream: std::ptr::null_mut(),
            image_type: 0,
            is_image_buffer: 0,
            image_buffer: std::ptr::null_mut(),
            image_buffer_size: 0,
            width: 0.,
            height: 0.,
            extension: std::ptr::null_mut(),
            x_dpi: 0.,
            y_dpi: 0.,
            chart: std::ptr::null_mut(),
            list_pointers: libxlsxwriter_sys::lxw_image_options__bindgen_ty_1 {
                stqe_next: std::ptr::null_mut(),
            },
        }
    }
}

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

#[derive(Debug, Clone, PartialEq, PartialOrd)]
pub struct HeaderFooterOptions {
    margin: f64,
}

impl Into<libxlsxwriter_sys::lxw_header_footer_options> for &HeaderFooterOptions {
    fn into(self) -> libxlsxwriter_sys::lxw_header_footer_options {
        libxlsxwriter_sys::lxw_header_footer_options {
            margin: self.margin,
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

impl Into<libxlsxwriter_sys::lxw_protection> for &Protection {
    fn into(self) -> libxlsxwriter_sys::lxw_protection {
        libxlsxwriter_sys::lxw_protection {
            no_select_locked_cells: convert_bool(self.no_select_locked_cells),
            no_select_unlocked_cells: convert_bool(self.no_select_unlocked_cells),
            format_cells: convert_bool(self.format_cells),
            format_columns: convert_bool(self.format_columns),
            format_rows: convert_bool(self.format_rows),
            insert_columns: convert_bool(self.insert_columns),
            insert_rows: convert_bool(self.insert_rows),
            insert_hyperlinks: convert_bool(self.insert_hyperlinks),
            delete_columns: convert_bool(self.delete_columns),
            delete_rows: convert_bool(self.delete_rows),
            sort: convert_bool(self.sort),
            autofilter: convert_bool(self.autofilter),
            pivot_tables: convert_bool(self.pivot_tables),
            scenarios: convert_bool(self.scenarios),
            objects: convert_bool(self.objects),
            no_content: convert_bool(self.no_content),
            no_objects: convert_bool(self.no_objects),
            no_sheet: convert_bool(false),
            is_configured: convert_bool(false),
            hash: [0; 5],
        }
    }
}

pub type WorksheetCol = libxlsxwriter_sys::lxw_col_t;
pub type WorksheetRow = libxlsxwriter_sys::lxw_row_t;
pub type RowColOptions = libxlsxwriter_sys::lxw_row_col_options;

pub struct Worksheet<'a> {
    pub(crate) _workbook: &'a Workbook,
    pub(crate) worksheet: *mut libxlsxwriter_sys::lxw_worksheet,
}

impl<'a> Worksheet<'a> {
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
                format.map(|x| x.format).unwrap_or(std::ptr::null_mut()),
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    pub fn write_string(
        &mut self,
        row: WorksheetRow,
        col: WorksheetCol,
        text: &str,
        format: Option<&Format>,
    ) -> Result<(), XlsxError> {
        unsafe {
            let result = libxlsxwriter_sys::worksheet_write_string(
                self.worksheet,
                row,
                col,
                CString::new(text).unwrap().as_c_str().as_ptr(),
                format.map(|x| x.format).unwrap_or(std::ptr::null_mut()),
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    pub fn write_formula(
        &mut self,
        row: WorksheetRow,
        col: WorksheetCol,
        formula: &str,
        format: Option<&Format>,
    ) -> Result<(), XlsxError> {
        unsafe {
            let result = libxlsxwriter_sys::worksheet_write_formula(
                self.worksheet,
                row,
                col,
                CString::new(formula).unwrap().as_c_str().as_ptr(),
                format.map(|x| x.format).unwrap_or(std::ptr::null_mut()),
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    pub fn write_array_formula(
        &mut self,
        first_row: WorksheetRow,
        first_col: WorksheetCol,
        last_row: WorksheetRow,
        last_col: WorksheetCol,
        formula: &str,
        format: Option<&Format>,
    ) -> Result<(), XlsxError> {
        unsafe {
            let result = libxlsxwriter_sys::worksheet_write_array_formula(
                self.worksheet,
                first_row,
                first_col,
                last_row,
                last_col,
                CString::new(formula).unwrap().as_c_str().as_ptr(),
                format.map(|x| x.format).unwrap_or(std::ptr::null_mut()),
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

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
                format.map(|x| x.format).unwrap_or(std::ptr::null_mut()),
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

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
                CString::new(url).unwrap().as_c_str().as_ptr(),
                format.map(|x| x.format).unwrap_or(std::ptr::null_mut()),
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

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
                format.map(|x| x.format).unwrap_or(std::ptr::null_mut()),
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

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
                format.map(|x| x.format).unwrap_or(std::ptr::null_mut()),
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

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
                CString::new(formula).unwrap().as_c_str().as_ptr(),
                format.map(|x| x.format).unwrap_or(std::ptr::null_mut()),
                number,
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    pub fn write_rich_string(
        &mut self,
        row: WorksheetRow,
        col: WorksheetCol,
        text: &[(&str, &Format)],
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
                format: x.1.format,
                string: y.as_mut_ptr() as *mut i8,
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
                format.map(|x| x.format).unwrap_or(std::ptr::null_mut()),
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
                format.map(|x| x.format).unwrap_or(std::ptr::null_mut()),
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
        options: &mut RowColOptions,
    ) -> Result<(), XlsxError> {
        unsafe {
            let result = libxlsxwriter_sys::worksheet_set_row_opt(
                self.worksheet,
                row,
                height,
                format.map(|x| x.format).unwrap_or(std::ptr::null_mut()),
                options,
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
        height: f64,
        format: Option<&Format>,
    ) -> Result<(), XlsxError> {
        unsafe {
            let result = libxlsxwriter_sys::worksheet_set_column(
                self.worksheet,
                first_col,
                last_col,
                height,
                format.map(|x| x.format).unwrap_or(std::ptr::null_mut()),
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
        height: f64,
        format: Option<&Format>,
        options: &mut RowColOptions,
    ) -> Result<(), XlsxError> {
        unsafe {
            let result = libxlsxwriter_sys::worksheet_set_column_opt(
                self.worksheet,
                first_col,
                last_col,
                height,
                format.map(|x| x.format).unwrap_or(std::ptr::null_mut()),
                options,
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    pub fn set_insert_image(
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

    pub fn set_insert_image_opt(
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

    pub fn set_insert_image_buffer(
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

    pub fn set_insert_image_buffer_opt(
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
                CString::new(string).unwrap().as_c_str().as_ptr(),
                format.map(|x| x.format).unwrap_or(std::ptr::null_mut()),
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    pub fn autofilter(
        &mut self,
        first_row: WorksheetRow,
        first_col: WorksheetCol,
        last_row: WorksheetRow,
        last_col: WorksheetCol,
    ) -> Result<(), XlsxError> {
        unsafe {
            let result = libxlsxwriter_sys::worksheet_autofilter(
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

    pub fn data_validation_cell(
        &mut self,
        row: WorksheetRow,
        col: WorksheetCol,
        validation: &DataValidation,
    ) -> Result<(), XlsxError> {
        unsafe {
            let result = libxlsxwriter_sys::worksheet_data_validation_cell(
                self.worksheet,
                row,
                col,
                &mut validation.to_c_struct().data_validation,
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    pub fn data_validation_range(
        &mut self,
        first_row: WorksheetRow,
        first_col: WorksheetCol,
        last_row: WorksheetRow,
        last_col: WorksheetCol,
        validation: &DataValidation,
    ) -> Result<(), XlsxError> {
        unsafe {
            let result = libxlsxwriter_sys::worksheet_data_validation_range(
                self.worksheet,
                first_row,
                first_col,
                last_row,
                last_col,
                &mut validation.to_c_struct().data_validation,
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    pub fn activate(&mut self) {
        unsafe {
            libxlsxwriter_sys::worksheet_activate(self.worksheet);
        }
    }

    pub fn select(&mut self) {
        unsafe {
            libxlsxwriter_sys::worksheet_select(self.worksheet);
        }
    }

    pub fn hide(&mut self) {
        unsafe {
            libxlsxwriter_sys::worksheet_hide(self.worksheet);
        }
    }

    pub fn set_first_sheet(&mut self) {
        unsafe {
            libxlsxwriter_sys::worksheet_set_first_sheet(self.worksheet);
        }
    }

    pub fn freeze_panes(&mut self, row: WorksheetRow, col: WorksheetCol) {
        unsafe {
            libxlsxwriter_sys::worksheet_freeze_panes(self.worksheet, row, col);
        }
    }

    pub fn split_panes(&mut self, vertical: f64, horizontal: f64) {
        unsafe {
            libxlsxwriter_sys::worksheet_split_panes(self.worksheet, vertical, horizontal);
        }
    }

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
