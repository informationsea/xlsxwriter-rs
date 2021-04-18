use super::super::{convert_bool, FormatColor};
use super::constants::*;

/// Struct to represent a chart pattern.
#[derive(Copy, Clone, PartialEq, PartialOrd)]
pub struct ChartPattern {
    /// The pattern foreground color.
    pub fg_color: FormatColor,
    /// The pattern background color.
    pub bg_color: FormatColor,
    /// The pattern type.
    pub chart_pattern: ChartPatternType,
}

impl ChartPattern {
    pub fn new(fg_color: FormatColor, bg_color: FormatColor, pattern: ChartPatternType) -> Self {
        ChartPattern {
            fg_color,
            bg_color,
            chart_pattern: pattern,
        }
    }

    pub(crate) fn value(&self) -> libxlsxwriter_sys::lxw_chart_pattern {
        libxlsxwriter_sys::lxw_chart_pattern {
            fg_color: self.fg_color.value(),
            bg_color: self.bg_color.value(),
            type_: self.chart_pattern.value(),
        }
    }
}

/// Struct to represent a chart line.
#[derive(Copy, Clone, PartialEq, PartialOrd)]
pub struct ChartLine {
    /// The chart font color.
    pub color: FormatColor,
    /// Turn off/hide line. Set to `false` or `true`.
    pub none: bool,
    /// Width of the line in increments of 0.25. Default is 2.25.
    pub width: f32,
    /// The line dash type.
    pub dash_type: ChartDashType,
    /// Set the transparency of the line. 0 - 100. Default 0.
    pub transparency: u8,
}

impl ChartLine {
    pub fn new() -> Self {
        ChartLine::default()
    }

    pub(crate) fn value(&self) -> libxlsxwriter_sys::lxw_chart_line {
        libxlsxwriter_sys::lxw_chart_line {
            color: self.color.value(),
            none: convert_bool(self.none),
            width: self.width,
            dash_type: self.dash_type.value(),
            transparency: self.transparency,
        }
    }
}

impl Default for ChartLine {
    fn default() -> Self {
        ChartLine {
            color: FormatColor::Black,
            none: false,
            width: 2.25,
            dash_type: ChartDashType::Solid,
            transparency: 0,
        }
    }
}

/// Struct to represent a chart fill.
#[derive(Clone, PartialEq, PartialOrd)]
pub struct ChartFill {
    /// The chart font color.
    pub color: FormatColor,
    /// Turn off/hide line. Set to false or true.
    pub none: bool,
    /// Set the transparency of the fill. 0 - 100. Default 0.
    pub transparency: u8,
}

impl ChartFill {
    pub fn new() -> Self {
        ChartFill::default()
    }

    pub(crate) fn value(&self) -> libxlsxwriter_sys::lxw_chart_fill {
        libxlsxwriter_sys::lxw_chart_fill {
            color: self.color.value(),
            none: convert_bool(self.none),
            transparency: self.transparency,
        }
    }
}

impl Default for ChartFill {
    fn default() -> Self {
        ChartFill {
            color: FormatColor::Black,
            none: false,
            transparency: 0,
        }
    }
}
