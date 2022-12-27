use crate::format::FormatColor;
use crate::worksheet::conditional_format::*;
use crate::{CStringHelper, StringOrFloat};

/// The 2 Color Scale type is used to specify Excel's "2 Color Scale" style conditional format.
#[derive(Debug, Clone, PartialEq, PartialOrd)]
pub struct TwoColorScaleCriteria {
    /// The rule used for the minimum condition in Color Scale conditional formats.
    pub min_rule_type: ConditionalFormatRuleTypes,
    /// The rule used for the maximum condition in Color Scale conditional formats.
    pub max_rule_type: ConditionalFormatRuleTypes,
    /// The minimum value used for Color Scale conditional formats.
    pub min_value: StringOrFloat,
    /// The maximum value used for Color Scale conditional formats.
    pub max_value: StringOrFloat,
    /// The color used for the minimum Color Scale conditional format.
    pub min_color: FormatColor,
    /// The color used for the maximum Color Scale conditional format.
    pub max_color: FormatColor,
}

impl TwoColorScaleCriteria {
    pub(crate) fn into_internal_value(
        &self,
        c_string_helper: &mut CStringHelper,
        conditional_format: &mut libxlsxwriter_sys::lxw_conditional_format,
    ) -> Result<(), XlsxError> {
        conditional_format.type_ =
            libxlsxwriter_sys::lxw_conditional_format_types_LXW_CONDITIONAL_2_COLOR_SCALE as u8;
        set_min_value(conditional_format, &self.min_value, c_string_helper)?;
        set_max_value(conditional_format, &self.max_value, c_string_helper)?;
        conditional_format.min_rule_type = self.min_rule_type.into_internal_value();
        conditional_format.max_rule_type = self.max_rule_type.into_internal_value();
        conditional_format.min_color = self.min_color.value();
        conditional_format.max_color = self.max_color.value();
        Ok(())
    }
}

impl ConditionalFormat {
    /// Two color scale
    ///
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # use xlsxwriter::worksheet::conditional_format::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-conditional_format-two-color-scale.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # for i in 0..30 {
    /// #     for j in 0..5 {
    /// #         worksheet.write_number(i, j, i.into(), None)?;
    /// #     }
    /// # }
    /// worksheet.conditional_format_range(
    ///     0, 0, 29, 0,
    ///     &ConditionalFormat::two_color_scale(
    ///         ConditionalFormatRuleTypes::Minimum,
    ///         ConditionalFormatRuleTypes::Maximum,
    ///         0.,
    ///         0.,
    ///         FormatColor::Yellow,
    ///         FormatColor::Red,
    ///     )
    /// )?;
    /// # worksheet.conditional_format_range(
    /// #     0, 1, 29, 1,
    /// #     &ConditionalFormat::two_color_scale(
    /// #         ConditionalFormatRuleTypes::Number,
    /// #         ConditionalFormatRuleTypes::Number,
    /// #         10.,
    /// #         20.,
    /// #         FormatColor::Yellow,
    /// #         FormatColor::Red,
    /// #     )
    /// # )?;
    /// # worksheet.conditional_format_range(
    /// #     0, 2, 29, 2,
    /// #     &ConditionalFormat::two_color_scale(
    /// #         ConditionalFormatRuleTypes::Percent,
    /// #         ConditionalFormatRuleTypes::Percent,
    /// #         10.,
    /// #         80.,
    /// #         FormatColor::Yellow,
    /// #         FormatColor::Red,
    /// #     )
    /// # )?;
    /// # worksheet.conditional_format_range(
    /// #     0, 3, 29, 3,
    /// #     &ConditionalFormat::two_color_scale(
    /// #         ConditionalFormatRuleTypes::Percentile,
    /// #         ConditionalFormatRuleTypes::Percentile,
    /// #         10.,
    /// #         80.,
    /// #         FormatColor::Yellow,
    /// #         FormatColor::Red,
    /// #     )
    /// # )?;
    /// # worksheet.conditional_format_range(
    /// #     0, 4, 29, 4,
    /// #     &ConditionalFormat::two_color_scale(
    /// #         ConditionalFormatRuleTypes::Formula,
    /// #         ConditionalFormatRuleTypes::Formula,
    /// #         "=20",
    /// #         "=25",
    /// #         FormatColor::Yellow,
    /// #         FormatColor::Red,
    /// #     )
    /// # )?;
    /// # Ok(())
    /// # }
    /// ```    

    pub fn two_color_scale<V1: Into<StringOrFloat>, V2: Into<StringOrFloat>>(
        min_rule_type: ConditionalFormatRuleTypes,
        max_rule_type: ConditionalFormatRuleTypes,
        min_value: V1,
        max_value: V2,
        min_color: FormatColor,
        max_color: FormatColor,
    ) -> ConditionalFormat {
        ConditionalFormat::TwoColorScale(TwoColorScaleCriteria {
            min_rule_type,
            max_rule_type,
            min_value: min_value.into(),
            max_value: max_value.into(),
            min_color,
            max_color,
        })
    }
}
