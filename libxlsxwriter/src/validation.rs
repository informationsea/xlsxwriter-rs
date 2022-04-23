use super::{convert_bool, DateTime};
use std::ffi::CString;
use std::os::raw::c_char;

#[derive(Debug, Copy, Clone, PartialEq, PartialOrd, Eq, Ord)]
pub enum DataValidationType {
    Integer,
    IntegerFormula,
    Decimal,
    DecimalFormula,
    List,
    ListFormula,
    Date,
    DateFormula,
    Time,
    TimeFormula,
    Length,
    LengthFormula,
    CustomFormula,
    Any,
}

impl DataValidationType {
    fn value(self) -> u8 {
        let value = match self {
            DataValidationType::Integer => {
                libxlsxwriter_sys::lxw_validation_types_LXW_VALIDATION_TYPE_INTEGER
            }
            DataValidationType::IntegerFormula => {
                libxlsxwriter_sys::lxw_validation_types_LXW_VALIDATION_TYPE_INTEGER_FORMULA
            }
            DataValidationType::Decimal => {
                libxlsxwriter_sys::lxw_validation_types_LXW_VALIDATION_TYPE_DECIMAL
            }
            DataValidationType::DecimalFormula => {
                libxlsxwriter_sys::lxw_validation_types_LXW_VALIDATION_TYPE_DECIMAL_FORMULA
            }
            DataValidationType::List => {
                libxlsxwriter_sys::lxw_validation_types_LXW_VALIDATION_TYPE_LIST
            }
            DataValidationType::ListFormula => {
                libxlsxwriter_sys::lxw_validation_types_LXW_VALIDATION_TYPE_LIST_FORMULA
            }
            DataValidationType::Date => {
                libxlsxwriter_sys::lxw_validation_types_LXW_VALIDATION_TYPE_DATE
            }
            DataValidationType::DateFormula => {
                libxlsxwriter_sys::lxw_validation_types_LXW_VALIDATION_TYPE_DATE_FORMULA
            }
            DataValidationType::Time => {
                libxlsxwriter_sys::lxw_validation_types_LXW_VALIDATION_TYPE_TIME
            }
            DataValidationType::TimeFormula => {
                libxlsxwriter_sys::lxw_validation_types_LXW_VALIDATION_TYPE_TIME_FORMULA
            }
            DataValidationType::Length => {
                libxlsxwriter_sys::lxw_validation_types_LXW_VALIDATION_TYPE_LENGTH
            }
            DataValidationType::LengthFormula => {
                libxlsxwriter_sys::lxw_validation_types_LXW_VALIDATION_TYPE_LENGTH_FORMULA
            }
            DataValidationType::CustomFormula => {
                libxlsxwriter_sys::lxw_validation_types_LXW_VALIDATION_TYPE_CUSTOM_FORMULA
            }
            DataValidationType::Any => {
                libxlsxwriter_sys::lxw_validation_types_LXW_VALIDATION_TYPE_ANY
            }
        };
        value as u8
    }
}

#[derive(Debug, Copy, Clone, PartialEq, PartialOrd, Eq, Ord)]
pub enum DataValidationCriteria {
    Between,
    NotBetween,
    EqualTo,
    NotEqualTo,
    GreaterThan,
    LessThan,
    GreaterThanOrEqualTo,
    LessThanOrEqualTo,
}

impl DataValidationCriteria {
    fn value(self) -> u8 {
        let value = match self {
            DataValidationCriteria::Between => libxlsxwriter_sys::lxw_validation_criteria_LXW_VALIDATION_CRITERIA_BETWEEN,
            DataValidationCriteria::NotBetween => libxlsxwriter_sys::lxw_validation_criteria_LXW_VALIDATION_CRITERIA_NOT_BETWEEN,
            DataValidationCriteria::EqualTo => libxlsxwriter_sys::lxw_validation_criteria_LXW_VALIDATION_CRITERIA_EQUAL_TO,
            DataValidationCriteria::NotEqualTo => libxlsxwriter_sys::lxw_validation_criteria_LXW_VALIDATION_CRITERIA_NOT_EQUAL_TO,
            DataValidationCriteria::GreaterThan => libxlsxwriter_sys::lxw_validation_criteria_LXW_VALIDATION_CRITERIA_GREATER_THAN,
            DataValidationCriteria::LessThan => libxlsxwriter_sys::lxw_validation_criteria_LXW_VALIDATION_CRITERIA_LESS_THAN,
            DataValidationCriteria::GreaterThanOrEqualTo => libxlsxwriter_sys::lxw_validation_criteria_LXW_VALIDATION_CRITERIA_GREATER_THAN_OR_EQUAL_TO,
            DataValidationCriteria::LessThanOrEqualTo => libxlsxwriter_sys::lxw_validation_criteria_LXW_VALIDATION_CRITERIA_LESS_THAN_OR_EQUAL_TO,
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

/// Worksheet data validation options.
#[derive(Debug, Clone, PartialEq)]
pub struct DataValidation {
    /// Set the validation type.    
    pub validate: DataValidationType,
    /// Set the validation criteria type to select the data.
    pub criteria: DataValidationCriteria,
    /// Controls whether a data validation is not applied to blank data in the cell.
    pub ignore_blank: bool,
    /// This parameter is used to toggle on and off the 'Show input message when cell is selected' option in the Excel data validation dialog.
    ///  When the option is off an input message is not displayed even if it has been set using `input_message`.
    pub show_input: bool,
    /// This parameter is used to toggle on and off the 'Show error alert after invalid data is entered' option in the Excel data validation dialog.
    /// When the option is off an error message is not displayed even if it has been set using `error_message`.
    pub show_error: bool,
    /// This parameter is used to specify the type of error dialog that is displayed.
    pub error_type: DataValidationErrorType,
    /// This parameter is used to toggle on and off the 'In-cell dropdown' option in the Excel data validation dialog.
    /// When the option is on a dropdown list will be shown for list validations.
    pub dropdown: bool,
    /// This parameter is used to set the limiting value to which the criteria is applied using a whole or decimal number.
    pub value_number: f64,
    /// This parameter is used to set the limiting value to which the criteria is applied using a cell reference. It is valid for any of the _FORMULA validation types.
    pub value_formula: Option<String>,
    /// This parameter is used to set a list of strings for a drop down list. The `value_formula` parameter can also be used to specify a list
    /// from an Excel cell range. Note, the string list is restricted by Excel to 255 characters, including comma separators.
    pub value_list: Option<Vec<String>>,
    /// This parameter is used to set the limiting value to which the date or time criteria is applied using a DateTime struct.
    pub value_datetime: DateTime,
    /// This parameter is the same as value_number but for the minimum value when a BETWEEN criteria is used.
    pub minimum_number: f64,
    /// This parameter is the same as value_formula but for the minimum value when a BETWEEN criteria is used.
    pub minimum_formula: Option<String>,
    /// This parameter is the same as value_datetime but for the minimum value when a BETWEEN criteria is used.
    pub minimum_datetime: DateTime,
    /// This parameter is the same as value_number but for the maximum value when a BETWEEN criteria is used.
    pub maximum_number: f64,
    /// This parameter is the same as value_formula but for the maximum value when a BETWEEN criteria is used.
    pub maximum_formula: Option<String>,
    /// This parameter is the same as value_datetime but for the maximum value when a BETWEEN criteria is used.
    pub maximum_datetime: DateTime,
    /// The input_title parameter is used to set the title of the input message that is displayed when a cell is entered.
    /// It has no default value and is only displayed if the input message is displayed. See the `input_message` parameter below.
    ///
    /// The maximum title length is 32 characters.
    pub input_title: Option<String>,
    /// The input_message parameter is used to set the input message that is displayed when a cell is entered. It has no default value.
    ///
    /// The message can be split over several lines using newlines. The maximum message length is 255 characters.
    pub input_message: Option<String>,
    /// The error_title parameter is used to set the title of the error message that is displayed when the data validation criteria is not met.
    /// The default error title is 'Microsoft Excel'. The maximum title length is 32 characters.
    pub error_title: Option<String>,
    /// The error_message parameter is used to set the error message that is displayed when a cell is entered. The default error message is
    /// "The value you entered is not valid. A user has restricted values that can be entered into the cell".
    ///
    /// The message can be split over several lines using newlines. The maximum message length is 255 characters.
    pub error_message: Option<String>,
}

fn option_str_to_cstr_bytes(s: &Option<String>) -> Option<Vec<u8>> {
    s.as_ref().map(|x| {
        CString::new(x as &str)
            .unwrap()
            .into_bytes_with_nul()
            .to_vec()
    })
}

impl DataValidation {
    pub fn new(
        validate: DataValidationType,
        criteria: DataValidationCriteria,
        error_type: DataValidationErrorType,
    ) -> DataValidation {
        DataValidation {
            validate,
            criteria,
            ignore_blank: true,
            show_input: false,
            show_error: true,
            error_type,
            dropdown: false,
            value_number: 0.,
            value_formula: None,
            value_list: None,
            value_datetime: DateTime {
                year: 0,
                month: 0,
                day: 0,
                hour: 0,
                min: 0,
                second: 0.,
            },
            minimum_number: 0.,
            minimum_formula: None,
            minimum_datetime: DateTime {
                year: 0,
                month: 0,
                day: 0,
                hour: 0,
                min: 0,
                second: 0.,
            },
            maximum_number: 0.,
            maximum_formula: None,
            maximum_datetime: DateTime {
                year: 0,
                month: 0,
                day: 0,
                hour: 0,
                min: 0,
                second: 0.,
            },
            input_title: None,
            input_message: None,
            error_title: None,
            error_message: None,
        }
    }
    pub(crate) fn to_c_struct(&self) -> CDataValidation {
        let mut _value_formula = option_str_to_cstr_bytes(&self.value_formula);
        let mut _value_list: Option<Vec<Vec<u8>>> = self.value_list.as_ref().map(|x| {
            x.iter()
                .map(|y| {
                    CString::new(y as &str)
                        .unwrap()
                        .into_bytes_with_nul()
                        .to_vec()
                })
                .collect()
        });
        let mut _value_list_ptr: Option<Vec<*mut c_char>> = _value_list.as_mut().map(|x| {
            x.iter_mut()
                .map(|y| y.as_mut_ptr() as *mut c_char)
                .collect()
        });
        if let Some(l) = _value_list_ptr.as_mut() {
            l.push(std::ptr::null_mut());
        }
        let mut _minimum_formula = option_str_to_cstr_bytes(&self.minimum_formula);
        let mut _maximum_formula = option_str_to_cstr_bytes(&self.maximum_formula);
        let mut _input_title = option_str_to_cstr_bytes(&self.input_title);
        let mut _input_message = option_str_to_cstr_bytes(&self.input_message);
        let mut _error_title = option_str_to_cstr_bytes(&self.error_title);
        let mut _error_message = option_str_to_cstr_bytes(&self.error_message);

        CDataValidation {
            data_validation: libxlsxwriter_sys::lxw_data_validation {
                validate: self.validate.value(),
                criteria: self.criteria.value(),
                ignore_blank: convert_bool(self.ignore_blank),
                show_input: convert_bool(self.show_input),
                show_error: convert_bool(self.show_error),
                error_type: self.error_type.value(),
                dropdown: convert_bool(self.dropdown),
                value_number: self.value_number,
                value_formula: _value_formula
                    .as_mut()
                    .map(|x| x.as_mut_ptr())
                    .unwrap_or(std::ptr::null_mut()) as *mut c_char,
                value_list: _value_list_ptr
                    .as_mut()
                    .map(|x| x.as_mut_ptr())
                    .unwrap_or(std::ptr::null_mut()),
                value_datetime: (&self.value_datetime).into(),
                minimum_number: self.minimum_number,
                minimum_formula: _minimum_formula
                    .as_mut()
                    .map(|x| x.as_mut_ptr())
                    .unwrap_or(std::ptr::null_mut())
                    as *mut c_char,
                minimum_datetime: (&self.minimum_datetime).into(),
                maximum_number: self.maximum_number,
                maximum_formula: _maximum_formula
                    .as_mut()
                    .map(|x| x.as_mut_ptr())
                    .unwrap_or(std::ptr::null_mut())
                    as *mut c_char,
                maximum_datetime: (&self.maximum_datetime).into(),
                input_title: _input_title
                    .as_mut()
                    .map(|x| x.as_mut_ptr())
                    .unwrap_or(std::ptr::null_mut()) as *mut c_char,
                input_message: _input_message
                    .as_mut()
                    .map(|x| x.as_mut_ptr())
                    .unwrap_or(std::ptr::null_mut()) as *mut c_char,
                error_title: _error_title
                    .as_mut()
                    .map(|x| x.as_mut_ptr())
                    .unwrap_or(std::ptr::null_mut()) as *mut c_char,
                error_message: _error_message
                    .as_mut()
                    .map(|x| x.as_mut_ptr())
                    .unwrap_or(std::ptr::null_mut()) as *mut c_char,
            },

            _value_formula,
            _value_list,
            _value_list_ptr,
            _minimum_formula,
            _maximum_formula,
            _input_title,
            _input_message,
            _error_title,
            _error_message,
        }
    }
}

#[derive(Debug, Clone)]
pub(crate) struct CDataValidation {
    _value_formula: Option<Vec<u8>>,
    _value_list: Option<Vec<Vec<u8>>>,
    _value_list_ptr: Option<Vec<*mut c_char>>,
    _minimum_formula: Option<Vec<u8>>,
    _maximum_formula: Option<Vec<u8>>,
    _input_title: Option<Vec<u8>>,
    _input_message: Option<Vec<u8>>,
    _error_title: Option<Vec<u8>>,
    _error_message: Option<Vec<u8>>,

    pub(crate) data_validation: libxlsxwriter_sys::lxw_data_validation,
}
