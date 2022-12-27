use crate::{convert_bool, prelude::FormatColor, CStringHelper, StringOrFloat, XlsxError};

use super::{set_max_value, set_min_value, ConditionalFormat, ConditionalFormatRuleTypes};

/// Values used to set the bar direction of a conditional format data bar.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum ConditionalFormatBarDirection {
    /// Data bar direction is set by Excel based on the context of the data displayed.    
    Context,
    /// Data bar direction is from right to left.
    RightToLeft,
    /// Data bar direction is from left to right.
    LeftToRight,
}

impl Default for ConditionalFormatBarDirection {
    fn default() -> Self {
        ConditionalFormatBarDirection::Context
    }
}

impl ConditionalFormatBarDirection {
    pub(crate) fn into_internal_type(self) -> u8 {
        let val = match self {
            ConditionalFormatBarDirection::Context => libxlsxwriter_sys::lxw_conditional_format_bar_direction_LXW_CONDITIONAL_BAR_DIRECTION_CONTEXT,
            ConditionalFormatBarDirection::RightToLeft => libxlsxwriter_sys::lxw_conditional_format_bar_direction_LXW_CONDITIONAL_BAR_DIRECTION_RIGHT_TO_LEFT,
            ConditionalFormatBarDirection::LeftToRight => libxlsxwriter_sys::lxw_conditional_format_bar_direction_LXW_CONDITIONAL_BAR_DIRECTION_LEFT_TO_RIGHT,
        };
        val as u8
    }
}

/// Values used to set the position of the axis in a conditional format data bar.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum ConditionalBarAxisPosition {
    /// Data bar axis position is set by Excel based on the context of the data displayed.    
    Automatic,
    /// Data bar axis position is set at the midpoint.
    Midpoint,
    /// Data bar axis is turned off.
    None,
}

impl Default for ConditionalBarAxisPosition {
    fn default() -> Self {
        ConditionalBarAxisPosition::Automatic
    }
}

impl ConditionalBarAxisPosition {
    pub(crate) fn into_internal_type(self) -> u8 {
        let val = match self {
            ConditionalBarAxisPosition::Automatic => libxlsxwriter_sys::lxw_conditional_bar_axis_position_LXW_CONDITIONAL_BAR_AXIS_AUTOMATIC,
            ConditionalBarAxisPosition::Midpoint => libxlsxwriter_sys::lxw_conditional_bar_axis_position_LXW_CONDITIONAL_BAR_AXIS_MIDPOINT,
            ConditionalBarAxisPosition::None => libxlsxwriter_sys::lxw_conditional_bar_axis_position_LXW_CONDITIONAL_BAR_AXIS_NONE,
        };
        val as u8
    }
}

/// The Data Bar type is used to specify Excel's "Data Bar" style conditional format.
#[derive(Debug, Clone, PartialEq, PartialOrd)]
pub struct ConditionalDataBar {
    pub min_rule_type: ConditionalFormatRuleTypes,
    pub max_rule_type: ConditionalFormatRuleTypes,
    pub min_value: StringOrFloat,
    pub max_value: StringOrFloat,
    pub bar_only: bool,
    pub color: Option<FormatColor>,
    pub solid: bool,
    pub negative_color: Option<FormatColor>,
    pub negative_color_same: bool,
    pub border_color: Option<FormatColor>,
    pub negative_border_color: Option<FormatColor>,
    pub negative_border_color_same: bool,
    pub no_border: bool,
    pub direction: ConditionalFormatBarDirection,
    pub axis_position: ConditionalBarAxisPosition,
    pub axis_color: Option<FormatColor>,
}

impl Default for ConditionalDataBar {
    fn default() -> Self {
        ConditionalDataBar {
            min_rule_type: ConditionalFormatRuleTypes::Minimum,
            max_rule_type: ConditionalFormatRuleTypes::Maximum,
            min_value: StringOrFloat::Float(0.),
            max_value: StringOrFloat::Float(0.),
            bar_only: false,
            color: Some(FormatColor::Blue),
            solid: false,
            negative_color: Some(FormatColor::Red),
            negative_color_same: false,
            border_color: None,
            negative_border_color: Some(FormatColor::Red),
            negative_border_color_same: false,
            no_border: false,
            direction: ConditionalFormatBarDirection::Context,
            axis_position: ConditionalBarAxisPosition::Automatic,
            axis_color: None,
        }
    }
}

impl From<ConditionalDataBar> for ConditionalFormat {
    fn from(value: ConditionalDataBar) -> Self {
        ConditionalFormat::DataBar(value)
    }
}

impl ConditionalDataBar {
    pub(crate) fn into_internal_value(
        &self,
        conditional_format: &mut libxlsxwriter_sys::lxw_conditional_format,
        c_string_helper: &mut CStringHelper,
    ) -> Result<(), XlsxError> {
        conditional_format.type_ =
            libxlsxwriter_sys::lxw_conditional_format_types_LXW_CONDITIONAL_DATA_BAR as u8;
        conditional_format.min_rule_type = self.min_rule_type.into_internal_value();
        conditional_format.max_rule_type = self.max_rule_type.into_internal_value();
        set_min_value(conditional_format, &self.min_value, c_string_helper)?;
        set_max_value(conditional_format, &self.max_value, c_string_helper)?;
        conditional_format.bar_only = convert_bool(self.bar_only);
        conditional_format.bar_color = self.color.map(|x| x.value()).unwrap_or(0);
        conditional_format.bar_solid = convert_bool(self.solid);
        conditional_format.bar_negative_color = self.negative_color.map(|x| x.value()).unwrap_or(0);
        conditional_format.bar_negative_color_same = convert_bool(self.negative_color_same);
        conditional_format.bar_border_color = self.border_color.map(|x| x.value()).unwrap_or(0);
        conditional_format.bar_negative_border_color =
            self.negative_border_color.map(|x| x.value()).unwrap_or(0);
        conditional_format.bar_negative_border_color_same =
            convert_bool(self.negative_border_color_same);
        conditional_format.bar_no_border = convert_bool(self.no_border);
        conditional_format.bar_direction = self.direction.into_internal_type();
        conditional_format.bar_axis_position = self.axis_position.into_internal_type();
        conditional_format.bar_axis_color = self.axis_color.map(|x| x.value()).unwrap_or(0);
        Ok(())
    }

    pub fn new() -> Self {
        Self::default()
    }

    pub fn min_rule_type(&mut self, min_rule_type: ConditionalFormatRuleTypes) -> &mut Self {
        self.min_rule_type = min_rule_type;
        self
    }

    pub fn max_rule_type(&mut self, max_rule_type: ConditionalFormatRuleTypes) -> &mut Self {
        self.max_rule_type = max_rule_type;
        self
    }

    pub fn min_value<V: Into<StringOrFloat>>(&mut self, min_value: V) -> &mut Self {
        self.min_value = min_value.into();
        self
    }

    pub fn max_value<V: Into<StringOrFloat>>(&mut self, max_value: V) -> &mut Self {
        self.max_value = max_value.into();
        self
    }

    pub fn bar_only(&mut self, bar_only: bool) -> &mut Self {
        self.bar_only = bar_only;
        self
    }

    pub fn color(&mut self, color: Option<FormatColor>) -> &mut Self {
        self.color = color;
        self
    }

    pub fn solid(&mut self, solid: bool) -> &mut Self {
        self.solid = solid;
        self
    }

    pub fn negative_color(&mut self, negative_color: Option<FormatColor>) -> &mut Self {
        self.negative_color = negative_color;
        self
    }

    pub fn negative_color_same(&mut self, negative_color_same: bool) -> &mut Self {
        self.negative_color_same = negative_color_same;
        self
    }

    pub fn border_color(&mut self, border_color: Option<FormatColor>) -> &mut Self {
        self.border_color = border_color;
        self
    }

    pub fn negative_border_color(
        &mut self,
        negative_border_color: Option<FormatColor>,
    ) -> &mut Self {
        self.negative_border_color = negative_border_color;
        self
    }

    pub fn negative_border_color_same(&mut self, negative_border_color_same: bool) -> &mut Self {
        self.negative_border_color_same = negative_border_color_same;
        self
    }

    pub fn no_border(&mut self, no_border: bool) -> &mut Self {
        self.no_border = no_border;
        self
    }

    pub fn direction(&mut self, direction: ConditionalFormatBarDirection) -> &mut Self {
        self.direction = direction;
        self
    }

    pub fn axis_position(&mut self, axis_position: ConditionalBarAxisPosition) -> &mut Self {
        self.axis_position = axis_position;
        self
    }

    pub fn axis_color(&mut self, axis_color: Option<FormatColor>) -> &mut Self {
        self.axis_color = axis_color;
        self
    }
}

impl ConditionalFormat {
    /// Data Bar
    ///
    /// Example:
    ///
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # use xlsxwriter::worksheet::conditional_format::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-conditional_format-data_bar.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # for i in 0..30 {
    /// #     for j in 0..2 {
    /// #         let v: f64 = i.into();
    /// #         worksheet.write_number(i, j, v - 15., None)?;
    /// #     }
    /// # }
    /// worksheet.conditional_format_range(
    ///     0, 0, 29, 0,
    ///     &ConditionalFormat::data_bar(
    ///         &ConditionalDataBar::new(),
    ///     )
    /// )?;
    /// worksheet.conditional_format_range(
    ///     0, 1, 29, 1,
    ///     &ConditionalFormat::data_bar(
    ///         ConditionalDataBar::new()
    ///             .min_rule_type(ConditionalFormatRuleTypes::Number)
    ///             .min_value(-5.0)
    ///             .max_rule_type(ConditionalFormatRuleTypes::Percent)
    ///             .max_value(90.0)
    ///             .bar_only(true)
    ///             .color(Some(FormatColor::Green))
    ///             .solid(true)
    ///             .negative_color(None)
    ///             .negative_color_same(true)
    ///             .border_color(None)
    ///             .negative_border_color(None)
    ///             .negative_border_color_same(true)
    ///             .no_border(true)
    ///             .direction(ConditionalFormatBarDirection::RightToLeft)
    ///             .axis_position(ConditionalBarAxisPosition::Midpoint)
    ///             .axis_color(Some(FormatColor::Purple)),
    ///     )
    /// )?;
    /// # Ok(())
    /// # }
    /// ```    

    pub fn data_bar(data_bar: &ConditionalDataBar) -> ConditionalFormat {
        ConditionalFormat::DataBar(data_bar.clone())
    }
}

#[cfg(test)]
mod test {}
