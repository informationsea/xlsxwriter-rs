use super::Workbook;
use std::ffi::CString;

#[allow(clippy::unreadable_literal)]
#[derive(Copy, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum FormatColor {
    Black,
    Blue,
    Brown,
    Cyan,
    Gray,
    Green,
    Lime,
    Magenta,
    Navy,
    Orange,
    Purple,
    Red,
    Pink,
    Silver,
    White,
    Yellow,
    Custom(i32),
}

#[allow(clippy::unreadable_literal)]
impl FormatColor {
    pub fn value(self) -> i32 {
        match self {
            FormatColor::Black => 0x1000000,
            FormatColor::Blue => 0x0000FF,
            FormatColor::Brown => 0x800000,
            FormatColor::Cyan => 0x00FFFF,
            FormatColor::Gray => 0x808080,
            FormatColor::Green => 0x008000,
            FormatColor::Lime => 0x00FF00,
            FormatColor::Magenta => 0xFF00FF,
            FormatColor::Navy => 0x000080,
            FormatColor::Orange => 0xFF6600,
            FormatColor::Purple => 0x800080,
            FormatColor::Red => 0xFF0000,
            FormatColor::Pink => 0xFF00FF,
            FormatColor::Silver => 0xC0C0C0,
            FormatColor::White => 0xFFFFFF,
            FormatColor::Yellow => 0xFFFF00,
            FormatColor::Custom(x) => x,
        }
    }
}

#[derive(Copy, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum FormatUnderline {
    Single,
    Double,
    SingleAccounting,
    DoubleAccounting,
}

impl FormatUnderline {
    pub fn value(self) -> u8 {
        let value = match self {
            FormatUnderline::Single => {
                libxlsxwriter_sys::lxw_format_underlines_LXW_UNDERLINE_SINGLE
            }
            FormatUnderline::SingleAccounting => {
                libxlsxwriter_sys::lxw_format_underlines_LXW_UNDERLINE_SINGLE_ACCOUNTING
            }
            FormatUnderline::Double => {
                libxlsxwriter_sys::lxw_format_underlines_LXW_UNDERLINE_DOUBLE
            }
            FormatUnderline::DoubleAccounting => {
                libxlsxwriter_sys::lxw_format_underlines_LXW_UNDERLINE_DOUBLE_ACCOUNTING
            }
        };
        value as u8
    }
}

#[derive(Copy, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum FormatScript {
    SuperScript,
    SubScript,
}

impl FormatScript {
    pub fn value(self) -> u8 {
        let value = match self {
            FormatScript::SuperScript => libxlsxwriter_sys::lxw_format_scripts_LXW_FONT_SUPERSCRIPT,
            FormatScript::SubScript => libxlsxwriter_sys::lxw_format_scripts_LXW_FONT_SUBSCRIPT,
        };
        value as u8
    }
}

#[derive(Copy, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum FormatAlignment {
    None,
    Left,
    Right,
    Fill,
    Justify,
    CenterAcross,
    Distributed,
    VerticalTop,
    VerticalBottom,
    VerticalCenter,
    VerticalJustify,
    VerticalDistributed,
}

impl FormatAlignment {
    pub fn value(self) -> u8 {
        let value = match self {
            FormatAlignment::None => libxlsxwriter_sys::lxw_format_alignments_LXW_ALIGN_NONE,
            FormatAlignment::Left => libxlsxwriter_sys::lxw_format_alignments_LXW_ALIGN_LEFT,
            FormatAlignment::Right => libxlsxwriter_sys::lxw_format_alignments_LXW_ALIGN_RIGHT,
            FormatAlignment::Fill => libxlsxwriter_sys::lxw_format_alignments_LXW_ALIGN_FILL,
            FormatAlignment::Justify => libxlsxwriter_sys::lxw_format_alignments_LXW_ALIGN_JUSTIFY,
            FormatAlignment::CenterAcross => {
                libxlsxwriter_sys::lxw_format_alignments_LXW_ALIGN_CENTER_ACROSS
            }
            FormatAlignment::Distributed => {
                libxlsxwriter_sys::lxw_format_alignments_LXW_ALIGN_DISTRIBUTED
            }
            FormatAlignment::VerticalTop => {
                libxlsxwriter_sys::lxw_format_alignments_LXW_ALIGN_VERTICAL_TOP
            }
            FormatAlignment::VerticalBottom => {
                libxlsxwriter_sys::lxw_format_alignments_LXW_ALIGN_VERTICAL_BOTTOM
            }
            FormatAlignment::VerticalCenter => {
                libxlsxwriter_sys::lxw_format_alignments_LXW_ALIGN_VERTICAL_CENTER
            }
            FormatAlignment::VerticalJustify => {
                libxlsxwriter_sys::lxw_format_alignments_LXW_ALIGN_VERTICAL_JUSTIFY
            }
            FormatAlignment::VerticalDistributed => {
                libxlsxwriter_sys::lxw_format_alignments_LXW_ALIGN_VERTICAL_DISTRIBUTED
            }
        };
        value as u8
    }
}

#[derive(Copy, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum FormatPatterns {
    None,
    Solid,
    MediumGray,
    DarkGray,
    LightGray,
    DarkHorizontal,
    DarkVertical,
    DarkDown,
    DarkUp,
    DarkGrid,
    DarkTrellis,
    LightHorizontal,
    LightVertical,
    LightDown,
    LightUp,
    LightGrid,
    LightTrellis,
    Gray125,
    Gray0625,
}

impl FormatPatterns {
    pub fn value(self) -> u8 {
        let value = match self {
            FormatPatterns::None => libxlsxwriter_sys::lxw_format_patterns_LXW_PATTERN_NONE,
            FormatPatterns::Solid => libxlsxwriter_sys::lxw_format_patterns_LXW_PATTERN_SOLID,
            FormatPatterns::MediumGray => {
                libxlsxwriter_sys::lxw_format_patterns_LXW_PATTERN_MEDIUM_GRAY
            }
            FormatPatterns::DarkGray => {
                libxlsxwriter_sys::lxw_format_patterns_LXW_PATTERN_DARK_GRAY
            }
            FormatPatterns::LightGray => {
                libxlsxwriter_sys::lxw_format_patterns_LXW_PATTERN_LIGHT_GRAY
            }
            FormatPatterns::DarkHorizontal => {
                libxlsxwriter_sys::lxw_format_patterns_LXW_PATTERN_DARK_HORIZONTAL
            }
            FormatPatterns::DarkVertical => {
                libxlsxwriter_sys::lxw_format_patterns_LXW_PATTERN_DARK_VERTICAL
            }
            FormatPatterns::DarkDown => {
                libxlsxwriter_sys::lxw_format_patterns_LXW_PATTERN_DARK_DOWN
            }
            FormatPatterns::DarkUp => libxlsxwriter_sys::lxw_format_patterns_LXW_PATTERN_DARK_UP,
            FormatPatterns::DarkGrid => {
                libxlsxwriter_sys::lxw_format_patterns_LXW_PATTERN_DARK_GRID
            }
            FormatPatterns::DarkTrellis => {
                libxlsxwriter_sys::lxw_format_patterns_LXW_PATTERN_DARK_TRELLIS
            }
            FormatPatterns::LightHorizontal => {
                libxlsxwriter_sys::lxw_format_patterns_LXW_PATTERN_LIGHT_HORIZONTAL
            }
            FormatPatterns::LightVertical => {
                libxlsxwriter_sys::lxw_format_patterns_LXW_PATTERN_LIGHT_VERTICAL
            }
            FormatPatterns::LightDown => {
                libxlsxwriter_sys::lxw_format_patterns_LXW_PATTERN_LIGHT_DOWN
            }
            FormatPatterns::LightUp => libxlsxwriter_sys::lxw_format_patterns_LXW_PATTERN_LIGHT_UP,
            FormatPatterns::LightGrid => {
                libxlsxwriter_sys::lxw_format_patterns_LXW_PATTERN_LIGHT_GRID
            }
            FormatPatterns::LightTrellis => {
                libxlsxwriter_sys::lxw_format_patterns_LXW_PATTERN_LIGHT_TRELLIS
            }
            FormatPatterns::Gray125 => libxlsxwriter_sys::lxw_format_patterns_LXW_PATTERN_GRAY_125,
            FormatPatterns::Gray0625 => {
                libxlsxwriter_sys::lxw_format_patterns_LXW_PATTERN_GRAY_0625
            }
        };
        value as u8
    }
}

#[derive(Copy, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum FormatBorder {
    None,
    Thin,
    Medium,
    Dashed,
    Dotted,
    Thick,
    Double,
    Hair,
    MediumDashed,
    DashDot,
    MediumDashDot,
    DashDotDot,
    MediumDashDotDot,
    SlantDashDot,
}

impl FormatBorder {
    pub fn value(self) -> u8 {
        let value = match self {
            FormatBorder::None => libxlsxwriter_sys::lxw_format_borders_LXW_BORDER_NONE,
            FormatBorder::Thin => libxlsxwriter_sys::lxw_format_borders_LXW_BORDER_THIN,
            FormatBorder::Medium => libxlsxwriter_sys::lxw_format_borders_LXW_BORDER_MEDIUM,
            FormatBorder::Dashed => libxlsxwriter_sys::lxw_format_borders_LXW_BORDER_DASHED,
            FormatBorder::Dotted => libxlsxwriter_sys::lxw_format_borders_LXW_BORDER_DOTTED,
            FormatBorder::Thick => libxlsxwriter_sys::lxw_format_borders_LXW_BORDER_THICK,
            FormatBorder::Double => libxlsxwriter_sys::lxw_format_borders_LXW_BORDER_DOUBLE,
            FormatBorder::Hair => libxlsxwriter_sys::lxw_format_borders_LXW_BORDER_HAIR,
            FormatBorder::MediumDashed => {
                libxlsxwriter_sys::lxw_format_borders_LXW_BORDER_MEDIUM_DASHED
            }
            FormatBorder::DashDot => libxlsxwriter_sys::lxw_format_borders_LXW_BORDER_DASH_DOT,
            FormatBorder::MediumDashDot => {
                libxlsxwriter_sys::lxw_format_borders_LXW_BORDER_MEDIUM_DASH_DOT
            }
            FormatBorder::DashDotDot => {
                libxlsxwriter_sys::lxw_format_borders_LXW_BORDER_DASH_DOT_DOT
            }
            FormatBorder::MediumDashDotDot => {
                libxlsxwriter_sys::lxw_format_borders_LXW_BORDER_MEDIUM_DASH_DOT_DOT
            }
            FormatBorder::SlantDashDot => {
                libxlsxwriter_sys::lxw_format_borders_LXW_BORDER_SLANT_DASH_DOT
            }
        };
        value as u8
    }
}

pub struct Format<'a> {
    pub(crate) _workbook: &'a Workbook,
    pub(crate) format: *mut libxlsxwriter_sys::lxw_format,
}

impl<'a> Format<'a> {
    pub fn set_font_name(&mut self, font_name: &str) {
        unsafe {
            libxlsxwriter_sys::format_set_font_name(
                self.format,
                CString::new(font_name).unwrap().as_c_str().as_ptr(),
            )
        }
    }

    pub fn set_font_size(&mut self, font_size: f64) {
        unsafe { libxlsxwriter_sys::format_set_font_size(self.format, font_size) }
    }

    pub fn set_font_color(&mut self, font_color: FormatColor) {
        unsafe { libxlsxwriter_sys::format_set_font_color(self.format, font_color.value()) }
    }

    pub fn set_bold(&mut self) {
        unsafe { libxlsxwriter_sys::format_set_bold(self.format) }
    }

    pub fn set_italic(&mut self) {
        unsafe { libxlsxwriter_sys::format_set_italic(self.format) }
    }

    pub fn set_underline(&mut self, underline: FormatUnderline) {
        unsafe { libxlsxwriter_sys::format_set_underline(self.format, underline.value()) }
    }

    pub fn set_font_strikeout(&mut self) {
        unsafe { libxlsxwriter_sys::format_set_font_strikeout(self.format) }
    }

    pub fn set_font_script(&mut self, script: FormatScript) {
        unsafe { libxlsxwriter_sys::format_set_font_script(self.format, script.value()) }
    }

    pub fn set_num_format(&mut self, num_font: &str) {
        unsafe {
            libxlsxwriter_sys::format_set_num_format(
                self.format,
                CString::new(num_font).unwrap().as_c_str().as_ptr(),
            )
        }
    }

    pub fn set_font_unlocked(&mut self) {
        unsafe { libxlsxwriter_sys::format_set_unlocked(self.format) }
    }

    pub fn set_font_hidden(&mut self) {
        unsafe { libxlsxwriter_sys::format_set_hidden(self.format) }
    }

    pub fn set_align(&mut self, align: FormatAlignment) {
        unsafe { libxlsxwriter_sys::format_set_align(self.format, align.value()) }
    }

    pub fn set_text_wrap(&mut self) {
        unsafe { libxlsxwriter_sys::format_set_text_wrap(self.format) }
    }

    pub fn set_rotation(&mut self, angle: i16) {
        unsafe { libxlsxwriter_sys::format_set_rotation(self.format, angle) }
    }

    pub fn set_indent(&mut self, level: u8) {
        unsafe { libxlsxwriter_sys::format_set_indent(self.format, level) }
    }

    pub fn set_shrink(&mut self) {
        unsafe { libxlsxwriter_sys::format_set_shrink(self.format) }
    }

    pub fn set_pattern(&mut self, pattern: FormatPatterns) {
        unsafe { libxlsxwriter_sys::format_set_pattern(self.format, pattern.value()) }
    }

    pub fn set_bg_color(&mut self, color: FormatColor) {
        unsafe { libxlsxwriter_sys::format_set_bg_color(self.format, color.value()) }
    }

    pub fn set_fg_color(&mut self, color: FormatColor) {
        unsafe { libxlsxwriter_sys::format_set_fg_color(self.format, color.value()) }
    }

    pub fn set_border(&mut self, border: FormatBorder) {
        unsafe { libxlsxwriter_sys::format_set_border(self.format, border.value()) }
    }

    pub fn set_border_bottom(&mut self, border: FormatBorder) {
        unsafe { libxlsxwriter_sys::format_set_bottom(self.format, border.value()) }
    }

    pub fn set_border_top(&mut self, border: FormatBorder) {
        unsafe { libxlsxwriter_sys::format_set_top(self.format, border.value()) }
    }

    pub fn set_border_left(&mut self, border: FormatBorder) {
        unsafe { libxlsxwriter_sys::format_set_left(self.format, border.value()) }
    }

    pub fn set_border_right(&mut self, border: FormatBorder) {
        unsafe { libxlsxwriter_sys::format_set_right(self.format, border.value()) }
    }

    pub fn set_border_color(&mut self, color: FormatColor) {
        unsafe { libxlsxwriter_sys::format_set_border_color(self.format, color.value()) }
    }

    pub fn set_border_bottom_color(&mut self, color: FormatColor) {
        unsafe { libxlsxwriter_sys::format_set_bottom_color(self.format, color.value()) }
    }

    pub fn set_border_top_color(&mut self, color: FormatColor) {
        unsafe { libxlsxwriter_sys::format_set_top_color(self.format, color.value()) }
    }

    pub fn set_border_left_color(&mut self, color: FormatColor) {
        unsafe { libxlsxwriter_sys::format_set_left_color(self.format, color.value()) }
    }

    pub fn set_border_right_color(&mut self, color: FormatColor) {
        unsafe { libxlsxwriter_sys::format_set_right_color(self.format, color.value()) }
    }
}
