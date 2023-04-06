use crate::{
    convert_validation_bool, try_to_vec, CStringHelper, Worksheet, WorksheetCol, WorksheetRow,
    XlsxError,
};

use super::DateTime;
use std::ffi::CString;
use std::os::raw::c_char;

#[derive(Debug, Clone, PartialEq, PartialOrd)]
pub enum DataValidationType {
    Integer{ignore_blank: bool, number_options: DataValidationNumberOptions<i64> },
    IntegerFormula{ignore_blank: bool, formula: String },
    Decimal{ignore_blank: bool, number_options: DataValidationNumberOptions<f64> },
    DecimalFormula{ignore_blank: bool, formula: String },
    List{ignore_blank: bool, dropdown: bool, values: Vec<String> },
    ListFormula{ignore_blank: bool, formula: String },
    Date{ignore_blank: bool, number_options: DataValidationNumberOptions<DateTime> },
    DateFormula{ignore_blank: bool, formula: String },
    Time{ignore_blank: bool, number_options: DataValidationNumberOptions<DateTime> },
    TimeFormula{ignore_blank: bool, formula: String },
    Length{ignore_blank: bool, number_options: DataValidationNumberOptions<usize> },
    LengthFormula{ignore_blank: bool, formula: String },
    CustomFormula{ignore_blank: bool, formula: String },
    Any,
}

pub struct DataValidation {
    input_message: Option<InputMessageOptions>,
    error_alert: Option<ErrorAlertOptions>,
    validation_type: DataValidationType
}

impl DataValidation {
    fn value(&self) -> u8 {
        let value = match self.validation_type {
            DataValidationType::Integer{..} => {
                libxlsxwriter_sys::lxw_validation_types_LXW_VALIDATION_TYPE_INTEGER
            }
            DataValidationType::IntegerFormula{..} => {
                libxlsxwriter_sys::lxw_validation_types_LXW_VALIDATION_TYPE_INTEGER_FORMULA
            }
            DataValidationType::Decimal{..} => {
                libxlsxwriter_sys::lxw_validation_types_LXW_VALIDATION_TYPE_DECIMAL
            }
            DataValidationType::DecimalFormula{..} => {
                libxlsxwriter_sys::lxw_validation_types_LXW_VALIDATION_TYPE_DECIMAL_FORMULA
            }
            DataValidationType::List{..} => {
                libxlsxwriter_sys::lxw_validation_types_LXW_VALIDATION_TYPE_LIST
            }
            DataValidationType::ListFormula{..} => {
                libxlsxwriter_sys::lxw_validation_types_LXW_VALIDATION_TYPE_LIST_FORMULA
            }
            DataValidationType::Date{..} => {
                libxlsxwriter_sys::lxw_validation_types_LXW_VALIDATION_TYPE_DATE
            }
            DataValidationType::DateFormula{..} => {
                libxlsxwriter_sys::lxw_validation_types_LXW_VALIDATION_TYPE_DATE_FORMULA
            }
            DataValidationType::Time{..} => {
                libxlsxwriter_sys::lxw_validation_types_LXW_VALIDATION_TYPE_TIME
            }
            DataValidationType::TimeFormula{..} => {
                libxlsxwriter_sys::lxw_validation_types_LXW_VALIDATION_TYPE_TIME_FORMULA
            }
            DataValidationType::Length{..} => {
                libxlsxwriter_sys::lxw_validation_types_LXW_VALIDATION_TYPE_LENGTH
            }
            DataValidationType::LengthFormula{..} => {
                libxlsxwriter_sys::lxw_validation_types_LXW_VALIDATION_TYPE_LENGTH_FORMULA
            }
            DataValidationType::CustomFormula{..} => {
                libxlsxwriter_sys::lxw_validation_types_LXW_VALIDATION_TYPE_CUSTOM_FORMULA
            }
            DataValidationType::Any{..} => {
                libxlsxwriter_sys::lxw_validation_types_LXW_VALIDATION_TYPE_ANY
            }
        };
        value as u8
    }
}

#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord)]
pub struct InputMessageOptions {
    pub title: String,
    pub message: String
}

#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord)]
pub struct ErrorAlertOptions {
    pub style: DataValidationErrorType,
    pub title: String,
    pub message: String
}

#[derive(Debug, Copy, Clone, PartialEq, PartialOrd)]
pub enum DataValidationNumberOptions<T> {
    Between(T, T),
    NotBetween(T, T),
    EqualTo(T),
    NotEqualTo(T),
    GreaterThan(T),
    LessThan(T),
    GreaterThanOrEqualTo(T),
    LessThanOrEqualTo(T),
}

impl<T> DataValidationNumberOptions<T> {
    fn value(&self) -> u8 {
        let value = match self {
            DataValidationNumberOptions::Between(_, _) => libxlsxwriter_sys::lxw_validation_criteria_LXW_VALIDATION_CRITERIA_BETWEEN,
            DataValidationNumberOptions::NotBetween(_, _) => libxlsxwriter_sys::lxw_validation_criteria_LXW_VALIDATION_CRITERIA_NOT_BETWEEN,
            DataValidationNumberOptions::EqualTo(_) => libxlsxwriter_sys::lxw_validation_criteria_LXW_VALIDATION_CRITERIA_EQUAL_TO,
            DataValidationNumberOptions::NotEqualTo(_) => libxlsxwriter_sys::lxw_validation_criteria_LXW_VALIDATION_CRITERIA_NOT_EQUAL_TO,
            DataValidationNumberOptions::GreaterThan(_) => libxlsxwriter_sys::lxw_validation_criteria_LXW_VALIDATION_CRITERIA_GREATER_THAN,
            DataValidationNumberOptions::LessThan(_) => libxlsxwriter_sys::lxw_validation_criteria_LXW_VALIDATION_CRITERIA_LESS_THAN,
            DataValidationNumberOptions::GreaterThanOrEqualTo(_) => libxlsxwriter_sys::lxw_validation_criteria_LXW_VALIDATION_CRITERIA_GREATER_THAN_OR_EQUAL_TO,
            DataValidationNumberOptions::LessThanOrEqualTo(_) => libxlsxwriter_sys::lxw_validation_criteria_LXW_VALIDATION_CRITERIA_LESS_THAN_OR_EQUAL_TO,
        };
        value as u8
    }
}

#[derive(Debug, Copy, Clone, PartialEq, PartialOrd, Eq, Ord)]
pub enum DataValidationErrorType {
    Stop,
    Warning,
    Information,
}

impl DataValidationErrorType {
    fn value(self) -> u8 {
        let value = match self {
            DataValidationErrorType::Stop => {
                libxlsxwriter_sys::lxw_validation_error_types_LXW_VALIDATION_ERROR_TYPE_STOP
            }
            DataValidationErrorType::Warning => {
                libxlsxwriter_sys::lxw_validation_error_types_LXW_VALIDATION_ERROR_TYPE_WARNING
            }
            DataValidationErrorType::Information => {
                libxlsxwriter_sys::lxw_validation_error_types_LXW_VALIDATION_ERROR_TYPE_INFORMATION
            }
        };
        value as u8
    }
}

impl DataValidation {
    pub fn new(validation_type: DataValidationType, input_message: Option<InputMessageOptions>, error_alert: Option<ErrorAlertOptions>) -> Self {
        DataValidation { input_message, error_alert, validation_type }
    }
    pub(crate) fn to_c_struct(
        &self,
        c_string_helper: &mut CStringHelper,
    ) -> Result<CDataValidation, XlsxError> {
        let mut _value_list: Option<Vec<Vec<u8>>> = match &self.validation_type {
            DataValidationType::List{values, ..} => {
                let mapped_vec = values.iter().map(|y| {
                CString::new(y as &str)
                    .unwrap()
                    .into_bytes_with_nul()
                    .to_vec()
                })
                .collect();
                Some(mapped_vec)},
            _ => None
        };
        let mut _value_list_ptr: Option<Vec<*mut c_char>> = match &self.validation_type {
            DataValidationType::List{values, ..} => Some(try_to_vec(values.iter().map(|y| Ok(c_string_helper.add(y)? as *mut c_char)))).transpose()?,
            _ => None
        };
        if let Some(l) = _value_list_ptr.as_mut() {
            l.push(std::ptr::null_mut());
        }
        let(minimum_number, maximum_number) = match self.validation_type {
            DataValidationType::Integer { number_options: 
                DataValidationNumberOptions::Between(x, y) | 
                DataValidationNumberOptions::NotBetween(x, y), .. } => (x as f64, y as f64), 
            DataValidationType::Decimal { number_options: 
                DataValidationNumberOptions::Between(x, y) | 
                DataValidationNumberOptions::NotBetween(x, y), .. } => (x, y), 
            DataValidationType::Length { number_options: 
                DataValidationNumberOptions::Between(x, y) | 
                DataValidationNumberOptions::NotBetween(x, y), .. } => (x as f64, y as f64), 
            _ => (0., 0.)
        };
        let (minimum_datetime, maximum_datetime) = match &self.validation_type {
            DataValidationType::Date { number_options: 
                DataValidationNumberOptions::Between(x, y) | 
                DataValidationNumberOptions::NotBetween(x, y), .. } => (x.into(), y.into()), 
            DataValidationType::Time { number_options: 
                DataValidationNumberOptions::Between(x, y) | 
                DataValidationNumberOptions::NotBetween(x, y), .. } => (x.into(), y.into()), 
            _ => ((&DateTime::default()).into(), (&DateTime::default()).into())
        };
        Ok(CDataValidation {
            data_validation: libxlsxwriter_sys::lxw_data_validation {
                validate: self.value(),
                criteria: match &self.validation_type {
                    DataValidationType::Date { number_options, ..} | DataValidationType::Time{ number_options, ..} => number_options.value(),
                    DataValidationType::Integer { number_options, .. } => number_options.value(),
                    DataValidationType::Length { number_options, .. }=> number_options.value(),
                    DataValidationType::Decimal { number_options, ..} => number_options.value(),
                    _ => 0u8,
                },
                ignore_blank: convert_validation_bool( match self.validation_type {
                    DataValidationType::Any => false,
                    DataValidationType::CustomFormula{ ignore_blank, ..} |
                    DataValidationType::Date { ignore_blank, ..} |
                    DataValidationType::DateFormula { ignore_blank, ..} |
                    DataValidationType::Decimal { ignore_blank, ..} |
                    DataValidationType::DecimalFormula { ignore_blank, .. } |
                    DataValidationType::Integer { ignore_blank, ..} |
                    DataValidationType::IntegerFormula { ignore_blank, ..} |
                    DataValidationType::Length { ignore_blank, ..} |
                    DataValidationType::LengthFormula { ignore_blank, .. } |
                    DataValidationType::List { ignore_blank, ..} |
                    DataValidationType::ListFormula { ignore_blank, .. } |
                    DataValidationType::Time { ignore_blank, ..} |
                    DataValidationType::TimeFormula { ignore_blank, .. } => ignore_blank
                }),
                show_input: convert_validation_bool( self.input_message.is_some()),
                show_error: convert_validation_bool(self.error_alert.is_some()),
                error_type: self.error_alert.as_ref().map(|x| x.style).unwrap_or(DataValidationErrorType::Stop).value(),
                dropdown: convert_validation_bool( match self.validation_type {
                    DataValidationType::List { dropdown, .. } => dropdown,
                    _ => true
                }),
                value_number: match self.validation_type {
                    DataValidationType::Integer { number_options: 
                        DataValidationNumberOptions::EqualTo(value) | 
                        DataValidationNumberOptions::NotEqualTo(value) |
                        DataValidationNumberOptions::GreaterThan(value) | 
                        DataValidationNumberOptions::GreaterThanOrEqualTo(value) | 
                        DataValidationNumberOptions::LessThan(value) |
                        DataValidationNumberOptions::LessThanOrEqualTo(value), 
                        ..
                    } => value as f64,
                    DataValidationType::Decimal { number_options: 
                        DataValidationNumberOptions::EqualTo(value) | 
                        DataValidationNumberOptions::NotEqualTo(value) |
                        DataValidationNumberOptions::GreaterThan(value) | 
                        DataValidationNumberOptions::GreaterThanOrEqualTo(value) | 
                        DataValidationNumberOptions::LessThan(value) |
                        DataValidationNumberOptions::LessThanOrEqualTo(value), 
                        ..
                    } => value,
                    DataValidationType::Length { number_options: 
                        DataValidationNumberOptions::EqualTo(value) | 
                        DataValidationNumberOptions::NotEqualTo(value) |
                        DataValidationNumberOptions::GreaterThan(value) | 
                        DataValidationNumberOptions::GreaterThanOrEqualTo(value) | 
                        DataValidationNumberOptions::LessThan(value) |
                        DataValidationNumberOptions::LessThanOrEqualTo(value), 
                        ..
                    } => value as f64,
                    _ => 0.
                },
                value_formula: c_string_helper.add_opt( match &self.validation_type {
                    DataValidationType::IntegerFormula{ formula, .. } |
                    DataValidationType::DecimalFormula{ formula, .. } |
                    DataValidationType::ListFormula{ formula, .. } |
                    DataValidationType::TimeFormula{ formula, .. } |
                    DataValidationType::DateFormula{ formula, .. } |
                    DataValidationType::LengthFormula{ formula, .. } |
                    DataValidationType::CustomFormula { formula, .. } => Some(formula),
                    _ => Some(""),
                })?
                    as *mut c_char,
                value_list: _value_list_ptr
                    .as_mut()
                    .map(|x| x.as_mut_ptr())
                    .unwrap_or(std::ptr::null_mut()),
                value_datetime: (&match &self.validation_type {
                        DataValidationType::Date { number_options: 
                            DataValidationNumberOptions::EqualTo(value) | 
                            DataValidationNumberOptions::NotEqualTo(value) |
                            DataValidationNumberOptions::GreaterThan(value) | 
                            DataValidationNumberOptions::GreaterThanOrEqualTo(value) | 
                            DataValidationNumberOptions::LessThan(value) |
                            DataValidationNumberOptions::LessThanOrEqualTo(value), 
                            .. 
                        } => value.clone(),
                        DataValidationType::Time { number_options: 
                            DataValidationNumberOptions::EqualTo(value) | 
                            DataValidationNumberOptions::NotEqualTo(value) |
                            DataValidationNumberOptions::GreaterThan(value) | 
                            DataValidationNumberOptions::GreaterThanOrEqualTo(value) | 
                            DataValidationNumberOptions::LessThan(value) |
                            DataValidationNumberOptions::LessThanOrEqualTo(value), 
                            .. 
                        } => value.clone(),
                        _ => DateTime::default()
                }).into(),
                minimum_number,
                minimum_formula: c_string_helper.add_opt(Some(""))?
                    as *mut c_char,
                minimum_datetime,
                maximum_number,
                maximum_formula: c_string_helper.add_opt(Some(""))?
                    as *mut c_char,
                maximum_datetime,
                input_title: c_string_helper.add_opt(self.input_message.clone().map(|x| x.title).as_deref())? as *mut c_char,
                input_message: c_string_helper.add_opt(self.input_message.clone().map(|x| x.message).as_deref())?
                    as *mut c_char,
                error_title: c_string_helper.add_opt(self.error_alert.clone().map(|x| x.title).as_deref())? as *mut c_char,
                error_message: c_string_helper.add_opt(self.error_alert.clone().map(|x| x.message).as_deref())?
                    as *mut c_char,
            },
            _value_list_ptr,
        })
    }
}

#[derive(Debug, Clone)]
pub(crate) struct CDataValidation {
    _value_list_ptr: Option<Vec<*mut c_char>>,
    pub(crate) data_validation: libxlsxwriter_sys::lxw_data_validation,
}

impl<'a> Worksheet<'a> {
    /// This function is used to construct an Excel data validation or to limit the user input to a dropdown list of values
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # use xlsxwriter::worksheet::validation::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-worksheet_validation-cell-3.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// let validation = DataValidation::new(
    ///     DataValidationType::Integer {
    ///         number_options: DataValidationNumberOptions::GreaterThanOrEqualTo(10),
    ///         ignore_blank: true
    ///     },
    ///     None,
    ///     Some(ErrorAlertOptions { 
    ///         style: DataValidationErrorType::Warning, 
    ///         title: String::new(), 
    ///         message: String::new()
    ///     })
    /// );
    ///
    /// worksheet.write_string(0, 0, "10 or greater", None)?;
    /// # worksheet.write_blank(1, 0, Some(&Format::new().set_border(FormatBorder::Dashed)))?;
    /// worksheet.data_validation_cell(1, 0, &validation)?;
    /// # workbook.close()
    /// # }
    /// ```    
    ///
    pub fn data_validation_cell(
        &mut self,
        row: WorksheetRow,
        col: WorksheetCol,
        validation: &DataValidation,
    ) -> Result<(), XlsxError> {
        unsafe {
            let mut c_string_helper = CStringHelper::new();
            let mut validation = validation.to_c_struct(&mut c_string_helper)?;
            let result = libxlsxwriter_sys::worksheet_data_validation_cell(
                self.worksheet,
                row,
                col,
                &mut validation.data_validation,
            );
            std::mem::drop(c_string_helper);
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    /// The this function is the same as the `data_validation_cell()`, see above, except the data validation is applied to a range of cells.
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # use xlsxwriter::worksheet::validation::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-worksheet_validation-cell-4.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// let validation = DataValidation::new(
    ///     DataValidationType::List {
    ///         values: vec!["VALUE1".to_string(), "VALUE2".to_string(), "VALUE3".to_string()],
    ///         dropdown: true,
    ///         ignore_blank: true
    ///     },
    ///     None,
    ///     None
    /// );
    ///
    /// # let format = workbook
    /// #    .add_format()
    /// #    .set_border(crate::FormatBorder::Dashed);
    /// #
    /// worksheet.data_validation_range(0, 0, 100, 100, &validation)?;
    /// # workbook.close()
    /// # }
    /// ```
    pub fn data_validation_range(
        &mut self,
        first_row: WorksheetRow,
        first_col: WorksheetCol,
        last_row: WorksheetRow,
        last_col: WorksheetCol,
        validation: &DataValidation,
    ) -> Result<(), XlsxError> {
        unsafe {
            let mut c_string_helper = CStringHelper::new();
            let result = libxlsxwriter_sys::worksheet_data_validation_range(
                self.worksheet,
                first_row,
                first_col,
                last_row,
                last_col,
                &mut validation
                    .to_c_struct(&mut c_string_helper)?
                    .data_validation,
            );
            std::mem::drop(c_string_helper);
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }
}

#[cfg(test)]
mod test {
    use super::*;
    use crate::{Format, FormatBorder, Workbook};

    #[test]
    fn test_validation() -> Result<(), XlsxError> {
        let workbook = Workbook::new("test-worksheet_validation-cell-1.xlsx")?;
        let validation = DataValidation::new(
            DataValidationType::Integer { 
                ignore_blank: true,
                number_options: DataValidationNumberOptions::Between(0, 2) 
            },
            Some(InputMessageOptions { 
                title: "Input Title".to_string(), 
                message: "Value must be 0 to 2".to_string()
            }),
            Some(ErrorAlertOptions { 
                style: DataValidationErrorType::Stop, 
                title: "Error Title".to_string(), 
                message: "Value must be 0 to 2".to_string()
            }) 
        );
        let mut worksheet = workbook.add_worksheet(None)?;
        worksheet.write_string(0, 0, "validation test", None)?;
        worksheet.write_blank(1, 0, Some(Format::new().set_border(FormatBorder::Dashed)))?;
        worksheet.data_validation_cell(1, 0, &validation)?;
        workbook.close()?;
        Ok(())
    }

    #[test]
    fn test_validation2() -> Result<(), XlsxError> {
        let workbook = Workbook::new("test-worksheet_validation-cell-2.xlsx")?;
        let validation = DataValidation::new(
            DataValidationType::List { 
                ignore_blank: true, 
                dropdown: true, 
                values: vec!["VALUE1".to_string(), "VALUE2".to_string()] 
            },
            Some(InputMessageOptions { 
                title: "Input Title".to_string(), 
                message: "Input Message".to_string()
            }),
            Some(ErrorAlertOptions { 
                style: DataValidationErrorType::Warning, 
                title: "Error Title".to_string(), 
                message: "Error Message".to_string()
            }), 
        );

        let mut worksheet = workbook.add_worksheet(None)?;
        worksheet.write_string(0, 0, "input list", None)?;
        for i in 1..=10 {
            for j in 0..=1 {
                worksheet.write_blank(
                    i,
                    j,
                    Some(Format::new().set_border(FormatBorder::Dashed)),
                )?;
            }
        }

        worksheet.data_validation_range(1, 0, 10, 1, &validation)?;
        workbook.close()?;
        Ok(())
    }
}
