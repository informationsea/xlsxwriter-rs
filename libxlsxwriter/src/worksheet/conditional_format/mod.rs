//! Conditional Format.
//!
//! Supported criteria.
//!
//! * [`ConditionalFormat::average`]
//! * [`ConditionalFormat::blanks`]
//! * [`ConditionalFormat::no_blanks`]
//! * [`ConditionalFormat::errors`]
//! * [`ConditionalFormat::no_errors`]
//! * [`ConditionalFormat::formula`]
//! * [`ConditionalFormat::duplicate`]
//! * [`ConditionalFormat::unique`]
//! * [`ConditionalFormat::top_num`]
//! * [`ConditionalFormat::top_percent`]
//! * [`ConditionalFormat::cell_between`]
//! * [`ConditionalFormat::cell_not_between`]
//! * [`ConditionalFormat::cell_equal_to`]
//! * [`ConditionalFormat::cell_not_equal_to`]
//! * [`ConditionalFormat::cell_greater_than`]
//! * [`ConditionalFormat::cell_greater_than_or_equal_to`]
//! * [`ConditionalFormat::cell_less_than`]
//! * [`ConditionalFormat::cell_less_than_or_equal_to`]
//! * [`ConditionalFormat::text_begins_with`]
//! * [`ConditionalFormat::text_ends_with`]
//! * [`ConditionalFormat::text_containing`]
//! * [`ConditionalFormat::text_not_containing`]
//! * [`ConditionalFormat::time_period`]
//! * [`ConditionalFormat::two_color_scale`]
//! * [`ConditionalFormat::three_color_scale`]
//! * [`ConditionalFormat::data_bar`]
//! * [`ConditionalFormat::icon_set`]

mod average;
mod cell;
mod data_bar;
mod icon;
mod text;
mod three_color;
mod time_period;
mod two_color;

pub use average::*;
pub use cell::*;
pub use data_bar::*;
pub use icon::*;
pub use text::*;
pub use three_color::*;
pub use time_period::*;
pub use two_color::*;

use crate::{
    convert_bool, CStringHelper, Format, StringOrFloat, Workbook, Worksheet, WorksheetCol,
    WorksheetRow, XlsxError,
};
use std::os::raw::c_char;

fn set_value_helper(
    float_value_store: &mut f64,
    string_value_store: &mut *mut c_char,
    value: &StringOrFloat,
    c_string_helper: &mut CStringHelper,
) -> Result<(), XlsxError> {
    match value {
        StringOrFloat::Float(f) => *float_value_store = *f,
        StringOrFloat::String(s) => *string_value_store = c_string_helper.add(s)? as *mut c_char,
    }
    Ok(())
}

fn set_value(
    conditional_format: &mut libxlsxwriter_sys::lxw_conditional_format,
    value: &StringOrFloat,
    c_string_helper: &mut CStringHelper,
) -> Result<(), XlsxError> {
    set_value_helper(
        &mut conditional_format.value,
        &mut conditional_format.value_string,
        value,
        c_string_helper,
    )
}

fn set_min_value(
    conditional_format: &mut libxlsxwriter_sys::lxw_conditional_format,
    value: &StringOrFloat,
    c_string_helper: &mut CStringHelper,
) -> Result<(), XlsxError> {
    set_value_helper(
        &mut conditional_format.min_value,
        &mut conditional_format.min_value_string,
        value,
        c_string_helper,
    )
}

fn set_max_value(
    conditional_format: &mut libxlsxwriter_sys::lxw_conditional_format,
    value: &StringOrFloat,
    c_string_helper: &mut CStringHelper,
) -> Result<(), XlsxError> {
    set_value_helper(
        &mut conditional_format.max_value,
        &mut conditional_format.max_value_string,
        value,
        c_string_helper,
    )
}

fn set_mid_value(
    conditional_format: &mut libxlsxwriter_sys::lxw_conditional_format,
    value: &StringOrFloat,
    c_string_helper: &mut CStringHelper,
) -> Result<(), XlsxError> {
    set_value_helper(
        &mut conditional_format.mid_value,
        &mut conditional_format.mid_value_string,
        value,
        c_string_helper,
    )
}

/// The Top or bottom type is used to specify the top n values by number or percentage in a range.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub enum TopOrBottomCriteria {
    TopOrBottomNum(u32),
    /// Format cells in the top of bottom percentage.
    TopOrBottomPercent(f64),
}

/// Conditional format rule types that apply to Color Scale and Data Bars.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum ConditionalFormatRuleTypes {
    /// Conditional format rule type: matches the minimum values in the range. Can only be applied to min_rule_type.
    Minimum,
    /// Conditional format rule type: use a number to set the bound.
    Number,
    /// Conditional format rule type: use a percentage to set the bound.
    Percent,
    /// Conditional format rule type: use a percentile to set the bound.
    Percentile,
    /// Conditional format rule type: use a formula to set the bound.
    Formula,
    /// Conditional format rule type: matches the maximum values in the range. Can only be applied to max_rule_type.
    Maximum,
}

impl ConditionalFormatRuleTypes {
    pub(crate) fn into_internal_value(self) -> u8 {
        let val = match self {
            ConditionalFormatRuleTypes::Minimum => libxlsxwriter_sys::lxw_conditional_format_rule_types_LXW_CONDITIONAL_RULE_TYPE_MINIMUM,
            ConditionalFormatRuleTypes::Number => libxlsxwriter_sys::lxw_conditional_format_rule_types_LXW_CONDITIONAL_RULE_TYPE_NUMBER,
            ConditionalFormatRuleTypes::Percent => libxlsxwriter_sys::lxw_conditional_format_rule_types_LXW_CONDITIONAL_RULE_TYPE_PERCENT,
            ConditionalFormatRuleTypes::Percentile => libxlsxwriter_sys::lxw_conditional_format_rule_types_LXW_CONDITIONAL_RULE_TYPE_PERCENTILE,
            ConditionalFormatRuleTypes::Formula => libxlsxwriter_sys::lxw_conditional_format_rule_types_LXW_CONDITIONAL_RULE_TYPE_FORMULA,
            ConditionalFormatRuleTypes::Maximum => libxlsxwriter_sys::lxw_conditional_format_rule_types_LXW_CONDITIONAL_RULE_TYPE_MAXIMUM,
        };
        val as u8
    }
}

/// See [`ConditionalFormat`] to learn more
#[derive(Debug, Clone, PartialEq, PartialOrd)]
pub enum ConditionalFormatTypes {
    /// The Cell type is the most common conditional formatting type. It is used when a format is applied to a cell based on a simple criterion.     
    Cell(ConditionalFormatCellCriteria),
    /// The Text type is used to specify Excel's "Specific Text" style conditional format.
    Text(ConditionalFormatTextCriteria),
    ///The Time Period type is used to specify Excel's "Dates Occurring" style conditional format.
    TimePeriod(ConditionalFormatTimePeriodCriteria),
    /// The Average type is used to specify Excel's "Average" style conditional format.
    Average(ConditionalFormatAverageCriteria),
    /// The Duplicate type is used to highlight duplicate cells in a range.
    Duplicate,
    /// The Unique type is used to highlight unique cells in a range.
    Unique,
    /// The Top type is used to specify the top n values by number or percentage in a range.
    Top(TopOrBottomCriteria),
    /// The Bottom type is used to specify the bottom n values by number or percentage in a range.
    Bottom(TopOrBottomCriteria),
    /// The Blanks type is used to highlight blank cells in a range.
    Blanks,
    /// The No Blanks type is used to highlight non blank cells in a range.
    NoBlanks,
    /// The Errors type is used to highlight error cells in a range.
    Errors,
    /// The No Errors type is used to highlight non error cells in a range.
    NoErrors,
    /// The Formula type is used to specify a conditional format based on a user defined formula.
    Formula(String),
}

impl Into<ConditionalFormatTypes> for ConditionalFormatCellCriteria {
    fn into(self) -> ConditionalFormatTypes {
        ConditionalFormatTypes::Cell(self)
    }
}

impl ConditionalFormatTypes {
    pub(crate) fn into_internal_value(
        &self,
        c_string_helper: &mut CStringHelper,
        conditional_format: &mut libxlsxwriter_sys::lxw_conditional_format,
    ) -> Result<(), XlsxError> {
        match self {
            ConditionalFormatTypes::Cell(criteria) => {
                criteria.into_internal_value(c_string_helper, conditional_format)?;
            }
            ConditionalFormatTypes::Text(criteria) => {
                criteria.into_internal_value(c_string_helper, conditional_format)?;
            }
            ConditionalFormatTypes::TimePeriod(criteria) => {
                criteria.into_internal_value(conditional_format)?;
            }
            ConditionalFormatTypes::Average(criteria) => {
                criteria.into_internal_value(conditional_format)?;
            }
            ConditionalFormatTypes::Duplicate => {
                conditional_format.type_ =
                    libxlsxwriter_sys::lxw_conditional_format_types_LXW_CONDITIONAL_TYPE_DUPLICATE
                        as u8;
            }
            ConditionalFormatTypes::Unique => {
                conditional_format.type_ =
                    libxlsxwriter_sys::lxw_conditional_format_types_LXW_CONDITIONAL_TYPE_UNIQUE
                        as u8;
            }
            ConditionalFormatTypes::Top(criteria) => {
                conditional_format.type_ =
                    libxlsxwriter_sys::lxw_conditional_format_types_LXW_CONDITIONAL_TYPE_TOP as u8;
                match criteria {
                    TopOrBottomCriteria::TopOrBottomNum(x) => {
                        conditional_format.criteria = 0;
                        conditional_format.min_value = (*x).into();
                        conditional_format.value = (*x).into();
                    }
                    TopOrBottomCriteria::TopOrBottomPercent(x) => {
                        conditional_format.criteria = libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_TOP_OR_BOTTOM_PERCENT as u8;
                        conditional_format.min_value = *x;
                        conditional_format.value = *x;
                    }
                }
            }
            ConditionalFormatTypes::Bottom(criteria) => {
                conditional_format.type_ =
                    libxlsxwriter_sys::lxw_conditional_format_types_LXW_CONDITIONAL_TYPE_BOTTOM
                        as u8;
                match criteria {
                    TopOrBottomCriteria::TopOrBottomNum(x) => {
                        conditional_format.criteria = 0;
                        conditional_format.min_value = (*x).into();
                    }
                    TopOrBottomCriteria::TopOrBottomPercent(x) => {
                        conditional_format.criteria = libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_TOP_OR_BOTTOM_PERCENT as u8;
                        conditional_format.min_value = *x;
                    }
                }
            }
            ConditionalFormatTypes::Blanks => {
                conditional_format.type_ =
                    libxlsxwriter_sys::lxw_conditional_format_types_LXW_CONDITIONAL_TYPE_BLANKS
                        as u8;
            }
            ConditionalFormatTypes::NoBlanks => {
                conditional_format.type_ =
                    libxlsxwriter_sys::lxw_conditional_format_types_LXW_CONDITIONAL_TYPE_NO_BLANKS
                        as u8;
            }
            ConditionalFormatTypes::Errors => {
                conditional_format.type_ =
                    libxlsxwriter_sys::lxw_conditional_format_types_LXW_CONDITIONAL_TYPE_ERRORS
                        as u8;
            }
            ConditionalFormatTypes::NoErrors => {
                conditional_format.type_ =
                    libxlsxwriter_sys::lxw_conditional_format_types_LXW_CONDITIONAL_TYPE_NO_ERRORS
                        as u8;
            }
            ConditionalFormatTypes::Formula(formula) => {
                conditional_format.type_ =
                    libxlsxwriter_sys::lxw_conditional_format_types_LXW_CONDITIONAL_TYPE_FORMULA
                        as u8;
                conditional_format.value_string = c_string_helper.add(formula)? as *mut c_char;
            }
        };
        Ok(())
    }
}

/// Conditional Format Criteria and Format.
///
/// Read methods' description in this enum to learn usage.
#[derive(Clone, PartialEq, PartialOrd)]
pub enum ConditionalFormat {
    ConditionType {
        criteria: ConditionalFormatTypes,
        format: Format,
    },
    TwoColorScale(TwoColorScaleCriteria),
    ThreeColorScale(ThreeColorScaleCriteria),
    DataBar(ConditionalDataBar),
    IconSet(ConditionalIconSet),
}

impl ConditionalFormat {
    /// Specify Excel's "Average" style conditional format.
    ///
    /// See [`ConditionalFormatAverageCriteria`] to learn available criteria.
    ///
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # use xlsxwriter::worksheet::conditional_format::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-conditional_format-average.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # for i in 0..=20 {
    /// #     worksheet.write_number(i, 0, i.into(), None)?;
    /// # }
    /// worksheet.conditional_format_range(
    ///     0, 0, 20, 0,
    ///     &ConditionalFormat::average(
    ///         ConditionalFormatAverageCriteria::AverageAbove,
    ///         Format::new().set_bg_color(FormatColor::Yellow),
    ///     ),
    /// )?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn average(
        average: ConditionalFormatAverageCriteria,
        format: &Format,
    ) -> ConditionalFormat {
        ConditionalFormat::ConditionType {
            criteria: ConditionalFormatTypes::Average(average),
            format: format.clone(),
        }
    }

    /// Highlight duplicate cells in a range.
    ///
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # use xlsxwriter::worksheet::conditional_format::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-conditional_format-duplicate.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # worksheet.write_string(0, 0, "A", None)?;
    /// # worksheet.write_string(1, 0, "A", None)?;
    /// # worksheet.write_string(2, 0, "B", None)?;
    /// worksheet.conditional_format_range(
    ///     0, 0, 2, 0,
    ///     &ConditionalFormat::duplicate(
    ///         Format::new().set_bg_color(FormatColor::Yellow)
    ///     )
    /// )?;
    /// # Ok(())
    /// # }
    /// ```    
    pub fn duplicate(format: &Format) -> ConditionalFormat {
        ConditionalFormat::ConditionType {
            criteria: ConditionalFormatTypes::Duplicate,
            format: format.clone(),
        }
    }

    /// Highlight unique cells in a range.
    ///
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # use xlsxwriter::worksheet::conditional_format::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-conditional_format-unique.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # worksheet.write_string(0, 0, "A", None)?;
    /// # worksheet.write_string(1, 0, "A", None)?;
    /// # worksheet.write_string(2, 0, "B", None)?;
    /// worksheet.conditional_format_range(
    ///     0, 0, 2, 0,
    ///     &ConditionalFormat::unique(
    ///         Format::new().set_bg_color(FormatColor::Yellow)
    ///     )
    /// )?;
    /// # Ok(())
    /// # }
    /// ```       
    pub fn unique(format: &Format) -> ConditionalFormat {
        ConditionalFormat::ConditionType {
            criteria: ConditionalFormatTypes::Unique,
            format: format.clone(),
        }
    }

    /// Highlight top N values cells in a range.
    ///
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # use xlsxwriter::worksheet::conditional_format::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-conditional_format-top-num.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # for i in 0..30 {
    /// #     worksheet.write_number(i, 0, i.into(), None)?;
    /// # }
    /// worksheet.conditional_format_range(
    ///     0, 0, 29, 0,
    ///     &ConditionalFormat::top_num(
    ///         5,
    ///         Format::new().set_bg_color(FormatColor::Yellow)
    ///     )
    /// )?;
    /// # Ok(())
    /// # }
    /// ```      
    pub fn top_num(num: u32, format: &Format) -> ConditionalFormat {
        ConditionalFormat::ConditionType {
            criteria: ConditionalFormatTypes::Top(TopOrBottomCriteria::TopOrBottomNum(num)),
            format: format.clone(),
        }
    }

    /// Highlight top N percent cells in a range.
    ///
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # use xlsxwriter::worksheet::conditional_format::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-conditional_format-top-percent.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # for i in 0..30 {
    /// #     worksheet.write_number(i, 0, i.into(), None)?;
    /// # }
    /// worksheet.conditional_format_range(
    ///     0, 0, 29, 0,
    ///     &ConditionalFormat::top_percent(
    ///         20.,
    ///         Format::new().set_bg_color(FormatColor::Yellow)
    ///     )
    /// )?;
    /// # Ok(())
    /// # }
    /// ```      
    pub fn top_percent(percent: f64, format: &Format) -> ConditionalFormat {
        ConditionalFormat::ConditionType {
            criteria: ConditionalFormatTypes::Top(TopOrBottomCriteria::TopOrBottomPercent(percent)),
            format: format.clone(),
        }
    }

    /// Highlight blank cells in a range.
    ///
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # use xlsxwriter::worksheet::conditional_format::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-conditional_format-blanks.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # worksheet.write_number(0, 0, 1.0, None)?;
    /// # worksheet.write_blank(1, 0, None)?;
    /// #
    /// worksheet.conditional_format_range(
    ///     0, 0, 1, 0,
    ///     &ConditionalFormat::blanks(
    ///         Format::new().set_bg_color(FormatColor::Yellow)
    ///     )
    /// )?;
    /// # Ok(())
    /// # }
    /// ```      
    pub fn blanks(format: &Format) -> ConditionalFormat {
        ConditionalFormat::ConditionType {
            criteria: ConditionalFormatTypes::Blanks,
            format: format.clone(),
        }
    }

    /// Highlight not blank cells in a range.
    ///
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # use xlsxwriter::worksheet::conditional_format::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-conditional_format-blanks.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # worksheet.write_number(0, 0, 1.0, None)?;
    /// # worksheet.write_blank(1, 0, None)?;
    /// #
    /// worksheet.conditional_format_range(
    ///     0, 0, 1, 0,
    ///     &ConditionalFormat::no_blanks(
    ///         Format::new().set_bg_color(FormatColor::Yellow)
    ///     )
    /// )?;
    /// # Ok(())
    /// # }
    /// ```      
    pub fn no_blanks(format: &Format) -> ConditionalFormat {
        ConditionalFormat::ConditionType {
            criteria: ConditionalFormatTypes::NoBlanks,
            format: format.clone(),
        }
    }

    /// Highlight error cells in a range.
    ///
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # use xlsxwriter::worksheet::conditional_format::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-conditional_format-error.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # worksheet.write_formula(0, 0, "=1+1", None)?;
    /// # worksheet.write_string(1, 1, "ABC", None)?;
    /// # worksheet.write_formula(1, 0, "=1+B2", None)?;
    /// #
    /// worksheet.conditional_format_range(
    ///     0, 0, 1, 0,
    ///     &ConditionalFormat::errors(
    ///         Format::new().set_bg_color(FormatColor::Yellow)
    ///     )
    /// )?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn errors(format: &Format) -> ConditionalFormat {
        ConditionalFormat::ConditionType {
            criteria: ConditionalFormatTypes::Errors,
            format: format.clone(),
        }
    }

    /// Highlight no error cells in a range.
    ///
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # use xlsxwriter::worksheet::conditional_format::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-conditional_format-no_error.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # worksheet.write_formula(0, 0, "=1+1", None)?;
    /// # worksheet.write_string(1, 1, "ABC", None)?;
    /// # worksheet.write_formula(1, 0, "=1+B2", None)?;
    /// #
    /// worksheet.conditional_format_range(
    ///     0, 0, 1, 0,
    ///     &ConditionalFormat::no_errors(
    ///         Format::new().set_bg_color(FormatColor::Yellow)
    ///     )
    /// )?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn no_errors(format: &Format) -> ConditionalFormat {
        ConditionalFormat::ConditionType {
            criteria: ConditionalFormatTypes::NoErrors,
            format: format.clone(),
        }
    }

    /// Highlight top N percent cells in a range.
    ///
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # use xlsxwriter::worksheet::conditional_format::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-conditional_format-formula.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # for i in 0..30 {
    /// #     worksheet.write_number(i, 0, i.into(), None)?;
    /// # }
    /// worksheet.conditional_format_range(
    ///     0, 0, 29, 0,
    ///     &ConditionalFormat::formula(
    ///         "=A1>10",
    ///         Format::new().set_bg_color(FormatColor::Yellow)
    ///     )
    /// )?;
    /// # Ok(())
    /// # }
    /// ```      
    pub fn formula(formula: &str, format: &Format) -> ConditionalFormat {
        ConditionalFormat::ConditionType {
            criteria: ConditionalFormatTypes::Formula(formula.to_string()),
            format: format.clone(),
        }
    }

    pub(crate) fn into_internal_type(
        &self,
        workbook: &Workbook,
        c_string_helper: &mut CStringHelper,
    ) -> Result<libxlsxwriter_sys::lxw_conditional_format, XlsxError> {
        let mut conditional_format = libxlsxwriter_sys::lxw_conditional_format {
            type_: 0,
            criteria: 0,
            value: 0.,
            value_string: std::ptr::null_mut(),
            format: std::ptr::null_mut(),
            min_value: 0.,
            min_value_string: std::ptr::null_mut(),
            min_rule_type: 0,
            min_color: 0,
            mid_value: 0.,
            mid_value_string: std::ptr::null_mut(),
            mid_rule_type: 0,
            mid_color: 0,
            max_value: 0.,
            max_value_string: std::ptr::null_mut(),
            max_rule_type: 0,
            max_color: 0,
            bar_color: 0,
            bar_only: convert_bool(false),
            data_bar_2010: 0,
            bar_solid: convert_bool(false),
            bar_negative_color: 0,
            bar_border_color: 0,
            bar_negative_border_color: 0,
            bar_negative_color_same: convert_bool(false),
            bar_negative_border_color_same: convert_bool(false),
            bar_no_border: convert_bool(false),
            bar_direction: 0,
            bar_axis_position: 0,
            bar_axis_color: 0,
            icon_style: 0,
            reverse_icons: convert_bool(false),
            icons_only: convert_bool(false),
            multi_range: std::ptr::null_mut(),
            stop_if_true: convert_bool(false),
        };
        match self {
            ConditionalFormat::ConditionType { criteria, format } => {
                let internal_format = workbook.get_internal_format(format)?;
                conditional_format.format = internal_format;
                criteria.into_internal_value(c_string_helper, &mut conditional_format)?;
            }
            ConditionalFormat::TwoColorScale(criteria) => {
                criteria.into_internal_value(c_string_helper, &mut conditional_format)?;
            }
            ConditionalFormat::ThreeColorScale(criteria) => {
                criteria.into_internal_value(c_string_helper, &mut conditional_format)?;
            }
            ConditionalFormat::DataBar(val) => {
                val.into_internal_value(&mut conditional_format, c_string_helper)?;
            }
            ConditionalFormat::IconSet(val) => {
                val.into_internal_value(&mut conditional_format)?;
            }
        }

        Ok(conditional_format)
    }
}

impl<'a> Worksheet<'a> {
    pub fn conditional_format_cell(
        &mut self,
        row: WorksheetRow,
        col: WorksheetCol,
        conditional_format: &ConditionalFormat,
    ) -> Result<(), XlsxError> {
        let mut c_string_helper = CStringHelper::new();
        unsafe {
            let mut conditional_format =
                conditional_format.into_internal_type(self._workbook, &mut c_string_helper)?;
            let result = libxlsxwriter_sys::worksheet_conditional_format_cell(
                self.worksheet,
                row,
                col,
                &mut conditional_format,
            );

            std::mem::drop(c_string_helper);
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    pub fn conditional_format_range(
        &mut self,
        first_row: WorksheetRow,
        first_col: WorksheetCol,
        last_row: WorksheetRow,
        last_col: WorksheetCol,
        conditional_format: &ConditionalFormat,
    ) -> Result<(), XlsxError> {
        let mut c_string_helper = CStringHelper::new();
        unsafe {
            let mut conditional_format =
                conditional_format.into_internal_type(self._workbook, &mut c_string_helper)?;
            let result = libxlsxwriter_sys::worksheet_conditional_format_range(
                self.worksheet,
                first_row,
                first_col,
                last_row,
                last_col,
                &mut conditional_format,
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
    use crate::*;

    #[test]
    fn test_worksheet_conditional_format4() -> Result<(), XlsxError> {
        let workbook = Workbook::new("test-worksheet_conditional-format4.xlsx")?;
        let mut worksheet = workbook.add_worksheet(Some("Conditional format"))?;
        for i in 0..=100 {
            for j in 0..10 {
                worksheet.write_number(i + 1, j, i.into(), None)?;
            }
        }
        let mut format = Format::new();
        format.set_bg_color(FormatColor::Yellow);

        worksheet.write_string(0, 0, "Top 10", None)?;
        worksheet.conditional_format_range(
            1,
            0,
            101,
            0,
            &ConditionalFormat::top_num(10, &format),
        )?;

        Ok(())
    }
}
