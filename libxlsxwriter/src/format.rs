use crate::{CStringHelper, XlsxError};

#[allow(clippy::unreadable_literal)]
#[derive(Copy, Clone, PartialEq, Eq, PartialOrd, Ord, Hash, Debug)]
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
    Custom(u32),
}

#[allow(clippy::unreadable_literal)]
impl FormatColor {
    pub fn value(self) -> u32 {
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

#[derive(Copy, Clone, PartialEq, Eq, PartialOrd, Ord, Hash, Debug)]
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

#[derive(Copy, Clone, PartialEq, Eq, PartialOrd, Ord, Hash, Debug)]
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

#[derive(Copy, Clone, PartialEq, Eq, PartialOrd, Ord, Hash, Debug)]
pub enum FormatAlignment {
    None,
    Left,
    Center,
    Right,
    Fill,
    Justify,
    CenterAcross,
    Distributed,
}

impl FormatAlignment {
    pub fn value(self) -> u8 {
        let value = match self {
            FormatAlignment::None => libxlsxwriter_sys::lxw_format_alignments_LXW_ALIGN_NONE,
            FormatAlignment::Left => libxlsxwriter_sys::lxw_format_alignments_LXW_ALIGN_LEFT,
            FormatAlignment::Center => libxlsxwriter_sys::lxw_format_alignments_LXW_ALIGN_CENTER,
            FormatAlignment::Right => libxlsxwriter_sys::lxw_format_alignments_LXW_ALIGN_RIGHT,
            FormatAlignment::Fill => libxlsxwriter_sys::lxw_format_alignments_LXW_ALIGN_FILL,
            FormatAlignment::Justify => libxlsxwriter_sys::lxw_format_alignments_LXW_ALIGN_JUSTIFY,
            FormatAlignment::CenterAcross => {
                libxlsxwriter_sys::lxw_format_alignments_LXW_ALIGN_CENTER_ACROSS
            }
            FormatAlignment::Distributed => {
                libxlsxwriter_sys::lxw_format_alignments_LXW_ALIGN_DISTRIBUTED
            }
        };
        value as u8
    }
}

#[derive(Copy, Clone, PartialEq, Eq, PartialOrd, Ord, Hash, Debug)]
pub enum FormatVerticalAlignment {
    None,
    VerticalTop,
    VerticalBottom,
    VerticalCenter,
    VerticalJustify,
    VerticalDistributed,
}

impl FormatVerticalAlignment {
    pub fn value(self) -> u8 {
        let value = match self {
            FormatVerticalAlignment::None => {
                libxlsxwriter_sys::lxw_format_alignments_LXW_ALIGN_NONE
            }
            FormatVerticalAlignment::VerticalTop => {
                libxlsxwriter_sys::lxw_format_alignments_LXW_ALIGN_VERTICAL_TOP
            }
            FormatVerticalAlignment::VerticalBottom => {
                libxlsxwriter_sys::lxw_format_alignments_LXW_ALIGN_VERTICAL_BOTTOM
            }
            FormatVerticalAlignment::VerticalCenter => {
                libxlsxwriter_sys::lxw_format_alignments_LXW_ALIGN_VERTICAL_CENTER
            }
            FormatVerticalAlignment::VerticalJustify => {
                libxlsxwriter_sys::lxw_format_alignments_LXW_ALIGN_VERTICAL_JUSTIFY
            }
            FormatVerticalAlignment::VerticalDistributed => {
                libxlsxwriter_sys::lxw_format_alignments_LXW_ALIGN_VERTICAL_DISTRIBUTED
            }
        };
        value as u8
    }
}

#[derive(Copy, Clone, PartialEq, Eq, PartialOrd, Ord, Hash, Debug)]
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

#[derive(Copy, Clone, PartialEq, Eq, PartialOrd, Ord, Hash, Debug)]
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

#[derive(Debug, Clone, PartialEq, PartialOrd, Eq, Hash, Default)]
pub struct Format {
    font_name: Option<String>,
    // font size / 100
    font_size: Option<u32>,
    font_color: Option<FormatColor>,
    bold: bool,
    italic: bool,
    underline: Option<FormatUnderline>,
    font_strikeout: bool,
    font_script: Option<FormatScript>,
    num_format: Option<String>,
    unlocked: bool,
    hidden: bool,
    align: Option<FormatAlignment>,
    vertical_align: Option<FormatVerticalAlignment>,
    rotation: Option<i16>,
    text_wrap: bool,
    indent: Option<u8>,
    shrink: bool,
    pattern: Option<FormatPatterns>,
    bg_color: Option<FormatColor>,
    fg_color: Option<FormatColor>,
    border: Option<FormatBorder>,
    bottom: Option<FormatBorder>,
    top: Option<FormatBorder>,
    left: Option<FormatBorder>,
    right: Option<FormatBorder>,
    border_color: Option<FormatColor>,
    bottom_color: Option<FormatColor>,
    top_color: Option<FormatColor>,
    left_color: Option<FormatColor>,
    right_color: Option<FormatColor>,
}

impl Format {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn set_font_name(&mut self, font_name: &str) -> &mut Self {
        self.font_name = Some(font_name.to_string());
        self
    }

    pub fn set_font_size(&mut self, font_size: f64) -> &mut Self {
        self.font_size = Some((font_size * 100.).round() as u32);
        self
    }

    pub fn set_font_color(&mut self, font_color: FormatColor) -> &mut Self {
        self.font_color = Some(font_color);
        self
    }

    pub fn set_bold(&mut self) -> &mut Self {
        self.bold = true;
        self
    }

    pub fn set_italic(&mut self) -> &mut Self {
        self.italic = true;
        self
    }

    pub fn set_underline(&mut self, underline: FormatUnderline) -> &mut Self {
        self.underline = Some(underline);
        self
    }

    pub fn set_font_strikeout(&mut self) -> &mut Self {
        self.font_strikeout = true;
        self
    }

    pub fn set_font_script(&mut self, script: FormatScript) -> &mut Self {
        self.font_script = Some(script);
        self
    }

    pub fn set_num_format(&mut self, num_format: &str) -> &mut Self {
        self.num_format = Some(num_format.to_string());
        self
    }

    pub fn set_unlocked(&mut self) -> &mut Self {
        self.unlocked = true;
        self
    }

    pub fn set_hidden(&mut self) -> &mut Self {
        self.hidden = true;
        self
    }

    pub fn set_align(&mut self, align: FormatAlignment) -> &mut Self {
        self.align = Some(align);
        self
    }

    pub fn set_vertical_align(&mut self, align: FormatVerticalAlignment) -> &mut Self {
        self.vertical_align = Some(align);
        self
    }

    pub fn set_text_wrap(&mut self) -> &mut Self {
        self.text_wrap = true;
        self
    }

    pub fn set_rotation(&mut self, angle: i16) -> &mut Self {
        self.rotation = Some(angle);
        self
    }

    pub fn set_indent(&mut self, level: u8) -> &mut Self {
        self.indent = Some(level);
        self
    }

    pub fn set_shrink(&mut self) -> &mut Self {
        self.shrink = true;
        self
    }

    pub fn set_pattern(&mut self, pattern: FormatPatterns) -> &mut Self {
        self.pattern = Some(pattern);
        self
    }

    pub fn set_bg_color(&mut self, color: FormatColor) -> &mut Self {
        self.bg_color = Some(color);
        self
    }

    pub fn set_fg_color(&mut self, color: FormatColor) -> &mut Self {
        self.fg_color = Some(color);
        self
    }

    pub fn set_border(&mut self, border: FormatBorder) -> &mut Self {
        self.border = Some(border);
        self
    }

    pub fn set_border_bottom(&mut self, border: FormatBorder) -> &mut Self {
        self.bottom = Some(border);
        self
    }

    pub fn set_border_top(&mut self, border: FormatBorder) -> &mut Self {
        self.top = Some(border);
        self
    }

    pub fn set_border_left(&mut self, border: FormatBorder) -> &mut Self {
        self.left = Some(border);
        self
    }

    pub fn set_border_right(&mut self, border: FormatBorder) -> &mut Self {
        self.right = Some(border);
        self
    }

    pub fn set_border_color(&mut self, color: FormatColor) -> &mut Self {
        self.border_color = Some(color);
        self
    }

    pub fn set_border_bottom_color(&mut self, color: FormatColor) -> &mut Self {
        self.bottom_color = Some(color);
        self
    }

    pub fn set_border_top_color(&mut self, color: FormatColor) -> &mut Self {
        self.top_color = Some(color);
        self
    }

    pub fn set_border_left_color(&mut self, color: FormatColor) -> &mut Self {
        self.left_color = Some(color);
        self
    }

    pub fn set_border_right_color(&mut self, color: FormatColor) -> &mut Self {
        self.right_color = Some(color);
        self
    }

    pub(crate) fn set_internal_format(
        &self,
        format: *mut libxlsxwriter_sys::lxw_format,
    ) -> Result<(), XlsxError> {
        let mut c_string_helper = CStringHelper::new();
        unsafe {
            if let Some(font_name) = self.font_name.as_deref() {
                libxlsxwriter_sys::format_set_font_name(format, c_string_helper.add(font_name)?);
            }

            if let Some(font_size) = self.font_size {
                let font_size: f64 = font_size.into();
                libxlsxwriter_sys::format_set_font_size(format, font_size / 100.0);
            }

            if let Some(font_color) = self.font_color {
                libxlsxwriter_sys::format_set_font_color(format, font_color.value());
            }

            if self.bold {
                libxlsxwriter_sys::format_set_bold(format);
            }

            if self.italic {
                libxlsxwriter_sys::format_set_italic(format);
            }

            if let Some(underline) = self.underline {
                libxlsxwriter_sys::format_set_underline(format, underline.value());
            }

            if self.font_strikeout {
                libxlsxwriter_sys::format_set_font_strikeout(format);
            }

            if let Some(font_script) = self.font_script {
                libxlsxwriter_sys::format_set_font_script(format, font_script.value());
            }

            if let Some(num_format) = self.num_format.as_deref() {
                libxlsxwriter_sys::format_set_num_format(format, c_string_helper.add(num_format)?);
            }

            if self.unlocked {
                libxlsxwriter_sys::format_set_unlocked(format);
            }

            if self.hidden {
                libxlsxwriter_sys::format_set_hidden(format);
            }

            if let Some(align) = self.align {
                libxlsxwriter_sys::format_set_align(format, align.value());
            }

            if let Some(vertical_align) = self.vertical_align {
                libxlsxwriter_sys::format_set_align(format, vertical_align.value());
            }

            if let Some(angle) = self.rotation {
                libxlsxwriter_sys::format_set_rotation(format, angle);
            }

            if self.text_wrap {
                libxlsxwriter_sys::format_set_text_wrap(format);
            }

            if let Some(indent) = self.indent {
                libxlsxwriter_sys::format_set_indent(format, indent);
            }

            if self.shrink {
                libxlsxwriter_sys::format_set_shrink(format);
            }

            if let Some(pattern) = self.pattern {
                libxlsxwriter_sys::format_set_pattern(format, pattern.value());
            }

            if let Some(bg_color) = self.bg_color {
                libxlsxwriter_sys::format_set_bg_color(format, bg_color.value());
            }

            if let Some(fg_color) = self.fg_color {
                libxlsxwriter_sys::format_set_bg_color(format, fg_color.value());
            }

            if let Some(style) = self.border {
                libxlsxwriter_sys::format_set_border(format, style.value());
            }

            if let Some(style) = self.bottom {
                libxlsxwriter_sys::format_set_bottom(format, style.value());
            }

            if let Some(style) = self.top {
                libxlsxwriter_sys::format_set_top(format, style.value());
            }

            if let Some(style) = self.left {
                libxlsxwriter_sys::format_set_left(format, style.value());
            }

            if let Some(style) = self.right {
                libxlsxwriter_sys::format_set_right(format, style.value());
            }

            if let Some(color) = self.border_color {
                libxlsxwriter_sys::format_set_border_color(format, color.value());
            }

            if let Some(color) = self.bottom_color {
                libxlsxwriter_sys::format_set_bottom_color(format, color.value());
            }

            if let Some(color) = self.top_color {
                libxlsxwriter_sys::format_set_top_color(format, color.value());
            }

            if let Some(color) = self.left_color {
                libxlsxwriter_sys::format_set_left_color(format, color.value());
            }

            if let Some(color) = self.right_color {
                libxlsxwriter_sys::format_set_right_color(format, color.value());
            }
        }
        Ok(())
    }
}
