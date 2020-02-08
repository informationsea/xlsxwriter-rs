use super::{convert_bool, DateTime};
use std::ffi::CString;

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

#[derive(Debug, Clone, PartialEq)]
pub struct DataValidation {
    pub validate: DataValidationType,
    pub criteria: DataValidationCriteria,
    pub ignore_blank: bool,
    pub show_input: bool,
    pub show_error: bool,
    pub error_type: DataValidationErrorType,
    pub dropdown: bool,
    pub value_number: f64,
    pub value_formula: Option<String>,
    pub value_list: Option<Vec<String>>,
    pub value_datetime: DateTime,
    pub minimum_number: f64,
    pub minimum_formula: Option<String>,
    pub minimum_datetime: DateTime,
    pub maximum_number: f64,
    pub maximum_formula: Option<String>,
    pub maximum_datetime: DateTime,
    pub input_title: Option<String>,
    pub input_message: Option<String>,
    pub error_title: Option<String>,
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
            show_input: true,
            show_error: true,
            error_type,
            dropdown: true,
            value_number: 0.,
            value_formula: None,
            value_list: None,
            value_datetime: DateTime {
                year: 2000,
                month: 1,
                day: 1,
                hour: 0,
                min: 0,
                second: 0.,
            },
            minimum_number: 0.,
            minimum_formula: None,
            minimum_datetime: DateTime {
                year: 2000,
                month: 1,
                day: 1,
                hour: 0,
                min: 0,
                second: 0.,
            },
            maximum_number: 0.,
            maximum_formula: None,
            maximum_datetime: DateTime {
                year: 2100,
                month: 1,
                day: 1,
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
        let mut value_formula = option_str_to_cstr_bytes(&self.value_formula);
        let mut value_list: Option<Vec<Vec<u8>>> = self.value_list.as_ref().map(|x| {
            x.iter()
                .map(|y| {
                    CString::new(y as &str)
                        .unwrap()
                        .into_bytes_with_nul()
                        .to_vec()
                })
                .collect()
        });
        let mut value_list_ptr: Option<Vec<*mut i8>> = value_list
            .as_mut()
            .map(|x| x.iter_mut().map(|y| y.as_mut_ptr() as *mut i8).collect());
        if let Some(l) = value_list_ptr.as_mut() {
            l.push(std::ptr::null_mut());
        }
        let mut minimum_formula = option_str_to_cstr_bytes(&self.minimum_formula);
        let mut maximum_formula = option_str_to_cstr_bytes(&self.maximum_formula);
        let mut input_title = option_str_to_cstr_bytes(&self.input_title);
        let mut input_message = option_str_to_cstr_bytes(&self.input_message);
        let mut error_title = option_str_to_cstr_bytes(&self.error_title);
        let mut error_message = option_str_to_cstr_bytes(&self.error_message);

        CDataValidation {
            data_validation: libxlsxwriter_sys::lxw_data_validation {
                validate: self.validate.value(),
                criteria: self.criteria.value(),
                ignore_blank: convert_bool(self.ignore_blank),
                show_input: convert_bool(self.show_input),
                show_error: convert_bool(self.show_error),
                error_type: self.error_type.value(),
                dropdown: convert_bool(self.dropdown),
                //is_between: convert_bool(false),
                value_number: self.value_number,
                value_formula: value_formula
                    .as_mut()
                    .map(|x| x.as_mut_ptr())
                    .unwrap_or(std::ptr::null_mut()) as *mut i8,
                value_list: value_list_ptr
                    .as_mut()
                    .map(|x| x.as_mut_ptr())
                    .unwrap_or(std::ptr::null_mut()),
                value_datetime: (&self.value_datetime).into(),
                minimum_number: self.minimum_number,
                minimum_formula: minimum_formula
                    .as_mut()
                    .map(|x| x.as_mut_ptr())
                    .unwrap_or(std::ptr::null_mut()) as *mut i8,
                minimum_datetime: (&self.minimum_datetime).into(),
                maximum_number: self.maximum_number,
                maximum_formula: maximum_formula
                    .as_mut()
                    .map(|x| x.as_mut_ptr())
                    .unwrap_or(std::ptr::null_mut()) as *mut i8,
                maximum_datetime: (&self.maximum_datetime).into(),
                input_title: input_title
                    .as_mut()
                    .map(|x| x.as_mut_ptr())
                    .unwrap_or(std::ptr::null_mut()) as *mut i8,
                input_message: input_message
                    .as_mut()
                    .map(|x| x.as_mut_ptr())
                    .unwrap_or(std::ptr::null_mut()) as *mut i8,
                error_title: error_title
                    .as_mut()
                    .map(|x| x.as_mut_ptr())
                    .unwrap_or(std::ptr::null_mut()) as *mut i8,
                error_message: error_message
                    .as_mut()
                    .map(|x| x.as_mut_ptr())
                    .unwrap_or(std::ptr::null_mut()) as *mut i8,
                //sqref: [0; 28],
                //list_pointers: libxlsxwriter_sys::lxw_data_val_obj__bindgen_ty_1 {
                //    stqe_next: std::ptr::null_mut(),
                //},
            },

            value_formula,
            value_list,
            value_list_ptr,
            minimum_formula,
            maximum_formula,
            input_title,
            input_message,
            error_title,
            error_message,
        }
    }
}

#[derive(Debug, Clone)]
pub(crate) struct CDataValidation {
    value_formula: Option<Vec<u8>>,
    value_list: Option<Vec<Vec<u8>>>,
    value_list_ptr: Option<Vec<*mut i8>>,
    minimum_formula: Option<Vec<u8>>,
    maximum_formula: Option<Vec<u8>>,
    input_title: Option<Vec<u8>>,
    input_message: Option<Vec<u8>>,
    error_title: Option<Vec<u8>>,
    error_message: Option<Vec<u8>>,

    pub(crate) data_validation: libxlsxwriter_sys::lxw_data_validation,
}
