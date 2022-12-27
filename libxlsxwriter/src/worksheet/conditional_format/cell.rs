use super::{set_max_value, set_min_value, set_value, ConditionalFormat, ConditionalFormatTypes};
use crate::{CStringHelper, Format, StringOrFloat, XlsxError};

/// The Cell type is the most common conditional formatting type. It is used when a format is
/// applied to a cell based on a simple criterion.
///
/// See [`ConditionalFormat`] to learn more.
#[derive(Debug, Clone, PartialEq, PartialOrd)]
pub enum ConditionalFormatCellCriteria {
    /// Format cells equal to a value.
    /// See usage at [`ConditionalFormat::cell_equal_to`]
    EqualTo(StringOrFloat),
    /// Format cells not equal to a value.
    /// See usage at [`ConditionalFormat::cell_not_equal_to`]
    NotEqualTo(StringOrFloat),
    /// Format cells greater than a value.
    GreaterThan(StringOrFloat),
    /// Format cells less than a value.
    LessThan(StringOrFloat),
    /// Format cells greater than or equal to a value.
    GreaterThanOrEqualTo(StringOrFloat),
    /// Format cells less than or equal to a value.
    LessThanOrEqualTo(StringOrFloat),
    /// Format cells between two values.
    Between {
        min: StringOrFloat,
        max: StringOrFloat,
    },
    /// Format cells that is not between two values.
    NotBetween {
        min: StringOrFloat,
        max: StringOrFloat,
    },
}

impl ConditionalFormatCellCriteria {
    pub(crate) fn into_internal_value(
        &self,
        c_string_helper: &mut CStringHelper,
        conditional_format: &mut libxlsxwriter_sys::lxw_conditional_format,
    ) -> Result<(), XlsxError> {
        conditional_format.type_ =
            libxlsxwriter_sys::lxw_conditional_format_types_LXW_CONDITIONAL_TYPE_CELL as u8;
        match self {
            ConditionalFormatCellCriteria::EqualTo(val) => {
                conditional_format.criteria =
                    libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_EQUAL_TO
                        as u8;
                set_value(conditional_format, val, c_string_helper)?;
            }
            ConditionalFormatCellCriteria::NotEqualTo(val) => {
                conditional_format.criteria =
                libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_NOT_EQUAL_TO
                    as u8;
                set_value(conditional_format, val, c_string_helper)?;
            }
            ConditionalFormatCellCriteria::GreaterThan(val) => {
                conditional_format.criteria =
                libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_GREATER_THAN
                    as u8;
                set_value(conditional_format, val, c_string_helper)?;
            }
            ConditionalFormatCellCriteria::LessThan(val) => {
                conditional_format.criteria =
                    libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_LESS_THAN
                        as u8;
                set_value(conditional_format, val, c_string_helper)?;
            }
            ConditionalFormatCellCriteria::GreaterThanOrEqualTo(val) => {
                conditional_format.criteria =
                libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_GREATER_THAN_OR_EQUAL_TO
                    as u8;
                set_value(conditional_format, val, c_string_helper)?;
            }
            ConditionalFormatCellCriteria::LessThanOrEqualTo(val) => {
                conditional_format.criteria =
                libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_LESS_THAN_OR_EQUAL_TO
                    as u8;
                set_value(conditional_format, val, c_string_helper)?;
            }
            ConditionalFormatCellCriteria::Between {
                min: min_val,
                max: max_val,
            } => {
                conditional_format.criteria =
                    libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_BETWEEN
                        as u8;
                set_min_value(conditional_format, min_val, c_string_helper)?;
                set_max_value(conditional_format, max_val, c_string_helper)?;
            }
            ConditionalFormatCellCriteria::NotBetween {
                min: min_val,
                max: max_val,
            } => {
                conditional_format.criteria =
                    libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_NOT_BETWEEN
                        as u8;
                set_min_value(conditional_format, min_val, c_string_helper)?;
                set_max_value(conditional_format, max_val, c_string_helper)?;
            }
        };
        Ok(())
    }
}

impl ConditionalFormat {
    /// Format cells equal to a value.
    ///
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # use xlsxwriter::worksheet::conditional_format::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-conditional_format-cell_equal_to.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # for i in 1..=10 {
    /// # worksheet.write_number(i, 0, i.into(), None)?;
    /// # }
    /// # worksheet.write_string(0, 0, "Equal to 3", None)?;
    /// worksheet.conditional_format_range(
    ///     1, 0, 10, 0,
    ///     &ConditionalFormat::cell_equal_to(3.0, Format::new().set_bg_color(FormatColor::Yellow)),
    /// )?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn cell_equal_to<V: Into<StringOrFloat>>(value: V, format: &Format) -> ConditionalFormat {
        ConditionalFormat::ConditionType {
            criteria: ConditionalFormatTypes::Cell(ConditionalFormatCellCriteria::EqualTo(
                value.into(),
            )),
            format: format.clone(),
        }
    }

    /// Format cells not equal to a value.
    ///
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # use xlsxwriter::worksheet::conditional_format::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-conditional_format-cell_not_equal_to.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # for i in 1..=10 {
    /// # worksheet.write_number(i, 0, i.into(), None)?;
    /// # }
    /// # worksheet.write_string(0, 0, "Not equal to 3", None)?;
    /// worksheet.conditional_format_range(
    ///     1, 0, 10, 0,
    ///     &ConditionalFormat::cell_not_equal_to(3.0, Format::new().set_bg_color(FormatColor::Yellow)),
    /// )?;
    /// # Ok(())
    /// # }
    /// ```    
    pub fn cell_not_equal_to<V: Into<StringOrFloat>>(
        value: V,
        format: &Format,
    ) -> ConditionalFormat {
        ConditionalFormat::ConditionType {
            criteria: ConditionalFormatTypes::Cell(ConditionalFormatCellCriteria::NotEqualTo(
                value.into(),
            )),
            format: format.clone(),
        }
    }

    /// Format cells greater than a value.
    ///
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # use xlsxwriter::worksheet::conditional_format::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-conditional_format-cell_greater_than.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # for i in 1..=10 {
    /// # worksheet.write_number(i, 0, i.into(), None)?;
    /// # }
    /// # worksheet.write_string(0, 0, "Greater than 3", None)?;
    /// worksheet.conditional_format_range(
    ///     1, 0, 10, 0,
    ///     &ConditionalFormat::cell_greater_than(3.0, Format::new().set_bg_color(FormatColor::Yellow)),
    /// )?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn cell_greater_than<V: Into<StringOrFloat>>(
        value: V,
        format: &Format,
    ) -> ConditionalFormat {
        ConditionalFormat::ConditionType {
            criteria: ConditionalFormatTypes::Cell(ConditionalFormatCellCriteria::GreaterThan(
                value.into(),
            )),
            format: format.clone(),
        }
    }

    /// Format cells less than a value.
    ///
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # use xlsxwriter::worksheet::conditional_format::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-conditional_format-cell_less_than.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # for i in 1..=10 {
    /// # worksheet.write_number(i, 0, i.into(), None)?;
    /// # }
    /// # worksheet.write_string(0, 0, "Less than 3", None)?;
    /// worksheet.conditional_format_range(
    ///     1, 0, 10, 0,
    ///     &ConditionalFormat::cell_less_than(3.0, Format::new().set_bg_color(FormatColor::Yellow)),
    /// )?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn cell_less_than<V: Into<StringOrFloat>>(value: V, format: &Format) -> ConditionalFormat {
        ConditionalFormat::ConditionType {
            criteria: ConditionalFormatTypes::Cell(ConditionalFormatCellCriteria::LessThan(
                value.into(),
            )),
            format: format.clone(),
        }
    }

    /// Format cells greater than or equal to a value.
    ///
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # use xlsxwriter::worksheet::conditional_format::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-conditional_format-cell_greater_than_or_equal_to.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # for i in 1..=10 {
    /// # worksheet.write_number(i, 0, i.into(), None)?;
    /// # }
    /// # worksheet.write_string(0, 0, "Greater than or equal to 3", None)?;
    /// worksheet.conditional_format_range(
    ///     1, 0, 10, 0,
    ///     &ConditionalFormat::cell_greater_than_or_equal_to(3.0, Format::new().set_bg_color(FormatColor::Yellow)),
    /// )?;
    /// # Ok(())
    /// # }
    /// ```    
    pub fn cell_greater_than_or_equal_to<V: Into<StringOrFloat>>(
        value: V,
        format: &Format,
    ) -> ConditionalFormat {
        ConditionalFormat::ConditionType {
            criteria: ConditionalFormatTypes::Cell(
                ConditionalFormatCellCriteria::GreaterThanOrEqualTo(value.into()),
            ),
            format: format.clone(),
        }
    }

    /// Format cells less than or equal to a value.
    ///
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # use xlsxwriter::worksheet::conditional_format::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-conditional_format-cell_less_than_or_equal_to.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # for i in 1..=10 {
    /// # worksheet.write_number(i, 0, i.into(), None)?;
    /// # }
    /// # worksheet.write_string(0, 0, "Less than or equal to 3", None)?;
    /// worksheet.conditional_format_range(
    ///     1, 0, 10, 0,
    ///     &ConditionalFormat::cell_less_than_or_equal_to(3.0, Format::new().set_bg_color(FormatColor::Yellow)),
    /// )?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn cell_less_than_or_equal_to<V: Into<StringOrFloat>>(
        value: V,
        format: &Format,
    ) -> ConditionalFormat {
        ConditionalFormat::ConditionType {
            criteria: ConditionalFormatTypes::Cell(
                ConditionalFormatCellCriteria::LessThanOrEqualTo(value.into()),
            ),
            format: format.clone(),
        }
    }

    /// Format cells between two values.
    ///
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # use xlsxwriter::worksheet::conditional_format::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-conditional_format-cell_between.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # for i in 1..=10 {
    /// # worksheet.write_number(i, 0, i.into(), None)?;
    /// # }
    /// # worksheet.write_string(0, 0, "Between from 3 to 5", None)?;
    /// worksheet.conditional_format_range(
    ///     1, 0, 10, 0,
    ///     &ConditionalFormat::cell_between(3.0, 5.0, Format::new().set_bg_color(FormatColor::Yellow)),
    /// )?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn cell_between<V1: Into<StringOrFloat>, V2: Into<StringOrFloat>>(
        min_value: V1,
        max_value: V2,
        format: &Format,
    ) -> ConditionalFormat {
        ConditionalFormat::ConditionType {
            criteria: ConditionalFormatTypes::Cell(ConditionalFormatCellCriteria::Between {
                min: min_value.into(),
                max: max_value.into(),
            }),
            format: format.clone(),
        }
    }

    /// Format cells not between two values.
    ///
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # use xlsxwriter::worksheet::conditional_format::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-conditional_format-cell_not_between.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # for i in 1..=10 {
    /// # worksheet.write_number(i, 0, i.into(), None)?;
    /// # }
    /// # worksheet.write_string(0, 0, "Not Between from 3 to 5", None)?;
    /// worksheet.conditional_format_range(
    ///     1, 0, 10, 0,
    ///     &ConditionalFormat::cell_not_between(3.0, 5.0, Format::new().set_bg_color(FormatColor::Yellow)),
    /// )?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn cell_not_between<V1: Into<StringOrFloat>, V2: Into<StringOrFloat>>(
        min_value: V1,
        max_value: V2,
        format: &Format,
    ) -> ConditionalFormat {
        ConditionalFormat::ConditionType {
            criteria: ConditionalFormatTypes::Cell(ConditionalFormatCellCriteria::NotBetween {
                min: min_value.into(),
                max: max_value.into(),
            }),
            format: format.clone(),
        }
    }
}

#[cfg(test)]
mod test {
    use crate::worksheet::conditional_format::*;
    use crate::*;

    #[test]
    fn test_worksheet_conditional_format1() -> Result<(), XlsxError> {
        let workbook = Workbook::new("test-worksheet_conditional-format1.xlsx")?;
        let mut worksheet = workbook.add_worksheet(Some("Conditional format"))?;

        for i in 0..=10 {
            for j in 0..=10 {
                worksheet.write_number(i + 1, j, i.into(), None)?;
            }
        }

        worksheet.write_string(0, 0, "Equal to \"3\"", None)?;
        worksheet.conditional_format_range(
            1,
            0,
            11,
            0,
            &ConditionalFormat::cell_equal_to(3.0, Format::new().set_bg_color(FormatColor::Yellow)),
        )?;

        worksheet.write_string(0, 1, "Not Equal to \"3\"", None)?;
        worksheet.conditional_format_range(
            1,
            1,
            11,
            1,
            &ConditionalFormat::cell_not_equal_to(
                3.0,
                Format::new().set_bg_color(FormatColor::Yellow),
            ),
        )?;

        worksheet.write_string(0, 2, "Greater than \"3\"", None)?;
        worksheet.conditional_format_range(
            1,
            2,
            11,
            2,
            &ConditionalFormat::cell_greater_than(
                3.0,
                Format::new().set_bg_color(FormatColor::Yellow),
            ),
        )?;

        worksheet.write_string(0, 3, "Less than \"3\"", None)?;
        worksheet.conditional_format_range(
            1,
            3,
            11,
            3,
            &ConditionalFormat::cell_less_than(
                3.0,
                Format::new().set_bg_color(FormatColor::Yellow),
            ),
        )?;

        // Cell Type

        workbook.close()?;
        Ok(())
    }
}
