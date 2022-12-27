use super::super::{convert_bool, FormatColor};
use super::{ChartDashType, ChartFill, ChartLine, ChartPattern, ChartPatternType};

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
