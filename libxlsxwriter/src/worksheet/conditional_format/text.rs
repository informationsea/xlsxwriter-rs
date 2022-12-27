use super::{ConditionalFormat, ConditionalFormatTypes};
use crate::{CStringHelper, Format, XlsxError};
use std::os::raw::c_char;

/// The Text type is used to specify Excel's "Specific Text" style conditional format.
///
/// See [`ConditionalFormat::text_containing`], [`ConditionalFormat::text_not_containing`],
/// [`ConditionalFormat::text_begins_with`], [`ConditionalFormat::text_ends_with`] to learn usage.
#[derive(Debug, Clone, PartialEq, PartialOrd)]
pub enum ConditionalFormatTextCriteria {
    /// Format cells that contain the specified text.     
    Containing(String),
    /// Format cells that don't contain the specified text.
    NotContaining(String),
    /// Format cells that begin with the specified text.
    BeginsWith(String),
    /// Format cells that end with the specified text.
    EndsWith(String),
}

impl ConditionalFormatTextCriteria {
    pub(crate) fn into_internal_value(
        &self,
        c_string_helper: &mut CStringHelper,
        conditional_format: &mut libxlsxwriter_sys::lxw_conditional_format,
    ) -> Result<(), XlsxError> {
        conditional_format.type_ =
            libxlsxwriter_sys::lxw_conditional_format_types_LXW_CONDITIONAL_TYPE_TEXT as u8;
        match self {
            ConditionalFormatTextCriteria::Containing(val) => {
                conditional_format.criteria = libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_TEXT_CONTAINING as u8;
                conditional_format.value_string = c_string_helper.add(val)? as *mut c_char;
            }
            ConditionalFormatTextCriteria::NotContaining(val) => {
                conditional_format.criteria = libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_TEXT_NOT_CONTAINING as u8;
                conditional_format.value_string = c_string_helper.add(val)? as *mut c_char;
            }
            ConditionalFormatTextCriteria::BeginsWith(val) => {
                conditional_format.criteria = libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_TEXT_BEGINS_WITH as u8;
                conditional_format.value_string = c_string_helper.add(val)? as *mut c_char;
            }
            ConditionalFormatTextCriteria::EndsWith(val) => {
                conditional_format.criteria = libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_TEXT_ENDS_WITH as u8;
                conditional_format.value_string = c_string_helper.add(val)? as *mut c_char;
            }
        }
        Ok(())
    }
}

impl ConditionalFormat {
    /// Format cells that contain the specified text.
    ///
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # use xlsxwriter::worksheet::conditional_format::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-conditional_format-text_contains.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # worksheet.write_string(1, 0, "Etiam", None)?;
    /// # worksheet.write_string(2, 0, "lobortis", None)?;
    /// # worksheet.write_string(3, 0, "ligula", None)?;
    /// # worksheet.write_string(4, 0, "eros", None)?;
    /// # worksheet.write_string(5, 0, "tincidunt", None)?;
    /// # worksheet.write_string(0, 0, "Containing ti", None)?;
    /// worksheet.conditional_format_range(
    ///     1, 0, 5, 0,
    ///     &ConditionalFormat::text_containing("ti", Format::new().set_bg_color(FormatColor::Yellow)),
    /// )?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn text_containing<T: ToString>(value: T, format: &Format) -> ConditionalFormat {
        ConditionalFormat::ConditionType {
            criteria: ConditionalFormatTypes::Text(ConditionalFormatTextCriteria::Containing(
                value.to_string(),
            )),
            format: format.clone(),
        }
    }

    /// Format cells that don't contain the specified text.
    ///
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # use xlsxwriter::worksheet::conditional_format::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-conditional_format-text_not_contains.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # worksheet.write_string(1, 0, "Etiam", None)?;
    /// # worksheet.write_string(2, 0, "lobortis", None)?;
    /// # worksheet.write_string(3, 0, "ligula", None)?;
    /// # worksheet.write_string(4, 0, "eros", None)?;
    /// # worksheet.write_string(5, 0, "tincidunt", None)?;
    /// # worksheet.write_string(0, 0, "Not Containing ti", None)?;
    /// worksheet.conditional_format_range(
    ///     1, 0, 5, 0,
    ///     &ConditionalFormat::text_not_containing("ti", Format::new().set_bg_color(FormatColor::Yellow)),
    /// )?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn text_not_containing<T: ToString>(value: T, format: &Format) -> ConditionalFormat {
        ConditionalFormat::ConditionType {
            criteria: ConditionalFormatTypes::Text(ConditionalFormatTextCriteria::NotContaining(
                value.to_string(),
            )),
            format: format.clone(),
        }
    }

    /// Format cells that begin with the specified text.
    ///
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # use xlsxwriter::worksheet::conditional_format::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-conditional_format-text_begins_with.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # worksheet.write_string(1, 0, "Etiam", None)?;
    /// # worksheet.write_string(2, 0, "lobortis", None)?;
    /// # worksheet.write_string(3, 0, "ligula", None)?;
    /// # worksheet.write_string(4, 0, "eros", None)?;
    /// # worksheet.write_string(5, 0, "tincidunt", None)?;
    /// # worksheet.write_string(0, 0, "Begins with li", None)?;
    /// worksheet.conditional_format_range(
    ///     1, 0, 5, 0,
    ///     &ConditionalFormat::text_begins_with("li", Format::new().set_bg_color(FormatColor::Yellow)),
    /// )?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn text_begins_with<T: ToString>(value: T, format: &Format) -> ConditionalFormat {
        ConditionalFormat::ConditionType {
            criteria: ConditionalFormatTypes::Text(ConditionalFormatTextCriteria::BeginsWith(
                value.to_string(),
            )),
            format: format.clone(),
        }
    }

    /// Format cells that end with the specified text.
    ///
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # use xlsxwriter::worksheet::conditional_format::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-conditional_format-text_ends_with.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # worksheet.write_string(1, 0, "Etiam", None)?;
    /// # worksheet.write_string(2, 0, "lobortis", None)?;
    /// # worksheet.write_string(3, 0, "ligula", None)?;
    /// # worksheet.write_string(4, 0, "eros", None)?;
    /// # worksheet.write_string(5, 0, "tincidunt", None)?;
    /// # worksheet.write_string(0, 0, "Ends with tis", None)?;
    /// worksheet.conditional_format_range(
    ///     1, 0, 5, 0,
    ///     &ConditionalFormat::text_ends_with("tis", Format::new().set_bg_color(FormatColor::Yellow)),
    /// )?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn text_ends_with<T: ToString>(value: T, format: &Format) -> ConditionalFormat {
        ConditionalFormat::ConditionType {
            criteria: ConditionalFormatTypes::Text(ConditionalFormatTextCriteria::EndsWith(
                value.to_string(),
            )),
            format: format.clone(),
        }
    }
}
