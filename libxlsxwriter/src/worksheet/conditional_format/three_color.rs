use crate::{prelude::*, CStringHelper};

use super::{
    set_max_value, set_mid_value, set_min_value, ConditionalFormat, ConditionalFormatRuleTypes,
};

/// The 3 Color Scale type is used to specify Excel's "3 Color Scale" style conditional format.
#[derive(Debug, Clone, PartialEq, PartialOrd)]
pub struct ThreeColorScaleCriteria {
    /// The rule used for the minimum condition in Color Scale conditional formats.    
    pub min_rule_type: ConditionalFormatRuleTypes,
    /// The rule used for the middle condition in Color Scale conditional formats.
    pub mid_rule_type: ConditionalFormatRuleTypes,
    /// The rule used for the maximum condition in Color Scale conditional formats.
    pub max_rule_type: ConditionalFormatRuleTypes,
    /// The minimum value used for Color Scale conditional formats.
    pub min_value: StringOrFloat,
    /// The middle value used for Color Scale conditional formats.
    pub mid_value: StringOrFloat,
    /// The maximum value used for Color Scale conditional formats.
    pub max_value: StringOrFloat,
    /// The color used for the minimum Color Scale conditional format.
    pub min_color: FormatColor,
    /// The color used for the middle Color Scale conditional format.
    pub mid_color: FormatColor,
    /// The color used for the maximum Color Scale conditional format.
    pub max_color: FormatColor,
}

impl ThreeColorScaleCriteria {
    pub(crate) fn into_internal_value(
        &self,
        c_string_helper: &mut CStringHelper,
        conditional_format: &mut libxlsxwriter_sys::lxw_conditional_format,
    ) -> Result<(), XlsxError> {
        conditional_format.type_ =
            libxlsxwriter_sys::lxw_conditional_format_types_LXW_CONDITIONAL_3_COLOR_SCALE as u8;
        set_min_value(conditional_format, &self.min_value, c_string_helper)?;
        set_mid_value(conditional_format, &self.mid_value, c_string_helper)?;
        set_max_value(conditional_format, &self.max_value, c_string_helper)?;
        conditional_format.min_rule_type = self.min_rule_type.into_internal_value();
        conditional_format.mid_rule_type = self.mid_rule_type.into_internal_value();
        conditional_format.max_rule_type = self.max_rule_type.into_internal_value();
        conditional_format.min_color = self.min_color.value();
        conditional_format.mid_color = self.mid_color.value();
        conditional_format.max_color = self.max_color.value();
        Ok(())
    }
}

impl ConditionalFormat {
    /// Three color scale
    ///
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # use xlsxwriter::worksheet::conditional_format::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-conditional_format-three-color-scale.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # for i in 0..30 {
    /// #     for j in 0..5 {
    /// #         worksheet.write_number(i, j, i.into(), None)?;
    /// #     }
    /// # }
    /// worksheet.conditional_format_range(
    ///     0, 0, 29, 0,
    ///     &ConditionalFormat::three_color_scale(
    ///         ConditionalFormatRuleTypes::Minimum,
    ///         ConditionalFormatRuleTypes::Percent,
    ///         ConditionalFormatRuleTypes::Maximum,
    ///         0.,
    ///         50.,
    ///         0.,
    ///         FormatColor::Yellow,
    ///         FormatColor::White,
    ///         FormatColor::Red,
    ///     )
    /// )?;
    /// # worksheet.conditional_format_range(
    /// #     0, 1, 29, 1,
    /// #     &ConditionalFormat::three_color_scale(
    /// #         ConditionalFormatRuleTypes::Number,
    /// #         ConditionalFormatRuleTypes::Number,
    /// #         ConditionalFormatRuleTypes::Number,
    /// #         5.,
    /// #         10.,
    /// #         15.,
    /// #         FormatColor::Yellow,
    /// #         FormatColor::White,
    /// #         FormatColor::Red,
    /// #     )
    /// # )?;
    /// # worksheet.conditional_format_range(
    /// #     0, 2, 29, 2,
    /// #     &ConditionalFormat::three_color_scale(
    /// #         ConditionalFormatRuleTypes::Percent,
    /// #         ConditionalFormatRuleTypes::Percent,
    /// #         ConditionalFormatRuleTypes::Percent,
    /// #         10.,
    /// #         20.,
    /// #         80.,
    /// #         FormatColor::Yellow,
    /// #         FormatColor::White,
    /// #         FormatColor::Red,
    /// #     )
    /// # )?;
    /// # worksheet.conditional_format_range(
    /// #     0, 3, 29, 3,
    /// #     &ConditionalFormat::three_color_scale(
    /// #         ConditionalFormatRuleTypes::Percentile,
    /// #         ConditionalFormatRuleTypes::Percentile,
    /// #         ConditionalFormatRuleTypes::Percentile,
    /// #         10.,
    /// #         70.,
    /// #         80.,
    /// #         FormatColor::Yellow,
    /// #         FormatColor::White,
    /// #         FormatColor::Red,
    /// #     )
    /// # )?;
    /// # worksheet.conditional_format_range(
    /// #     0, 4, 29, 4,
    /// #     &ConditionalFormat::three_color_scale(
    /// #         ConditionalFormatRuleTypes::Formula,
    /// #         ConditionalFormatRuleTypes::Formula,
    /// #         ConditionalFormatRuleTypes::Formula,
    /// #         "=20",
    /// #         "=23",
    /// #         "=25",
    /// #         FormatColor::Yellow,
    /// #         FormatColor::White,
    /// #         FormatColor::Red,
    /// #     )
    /// # )?;
    /// # Ok(())
    /// # }
    /// ```    

    pub fn three_color_scale<
        V1: Into<StringOrFloat>,
        V2: Into<StringOrFloat>,
        V3: Into<StringOrFloat>,
    >(
        min_rule_type: ConditionalFormatRuleTypes,
        mid_rule_type: ConditionalFormatRuleTypes,
        max_rule_type: ConditionalFormatRuleTypes,
        min_value: V1,
        mid_value: V2,
        max_value: V3,
        min_color: FormatColor,
        mid_color: FormatColor,
        max_color: FormatColor,
    ) -> ConditionalFormat {
        ConditionalFormat::ThreeColorScale(ThreeColorScaleCriteria {
            min_rule_type,
            mid_rule_type,
            max_rule_type,
            min_value: min_value.into(),
            mid_value: mid_value.into(),
            max_value: max_value.into(),
            min_color,
            mid_color,
            max_color,
        })
    }
}
