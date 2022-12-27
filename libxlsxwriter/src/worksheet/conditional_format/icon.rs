use crate::{convert_bool, XlsxError};

use super::ConditionalFormat;

/// Definitions of icon styles used with Icon Set conditional formats.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum ConditionalIconType {
    /// Icon style: 3 colored arrows showing up, sideways and down.
    Icons3ArrowColored,
    /// Icon style: 3 gray arrows showing up, sideways and down.
    Icons3ArrowGray,
    /// Icon style: 3 colored flags in red, yellow and green.
    Icons3Flags,
    /// Icon style: 3 traffic lights - rounded.
    Icons3TrafficLightsUnrimmed,
    /// Icon style: 3 traffic lights with a rim - squarish.
    Icons3TrafficLightsRimmed,
    /// Icon style: 3 colored shapes - a circle, triangle and diamond.
    Icons3Signs,
    /// Icon style: 3 circled symbols with tick mark, exclamation and cross.
    Icons3SymbolsCircled,
    /// Icon style: 3 symbols with tick mark, exclamation and cross.
    Icons3SymbolsUncircled,
    // Icon style: 4 colored arrows showing up, diagonal up, diagonal down and down.
    Icons4ArrowColored,
    /// Icon style: 4 gray arrows showing up, diagonal up, diagonal down and down.
    Icons4ArrowGray,
    /// Icon style: 4 circles in 4 colors going from red to black.
    Icons4RedToBlack,
    /// Icon style: 4 histogram ratings.
    Icons4Rating,
    /// Icon style: 4 traffic lights.
    Icons4TrafficLights,
    /// Icon style: 5 colored arrows showing up, diagonal up, sideways, diagonal down and down.
    Icons5ArrowColored,
    /// Icon style: 5 gray arrows showing up, diagonal up, sideways, diagonal down and down.
    Icons5ArrowGray,
    /// Icon style: 5 histogram ratings.
    Icons5Rating,
    /// Icon style: 5 quarters, from 0 to 4 quadrants filled.
    Icons5Quarters,
}

impl ConditionalIconType {
    pub(crate) fn into_internal_type(self) -> u8 {
        let val = match self {
            ConditionalIconType::Icons3ArrowColored => {
                libxlsxwriter_sys::lxw_conditional_icon_types_LXW_CONDITIONAL_ICONS_3_ARROWS_COLORED
            }
            ConditionalIconType::Icons3ArrowGray => {
                libxlsxwriter_sys::lxw_conditional_icon_types_LXW_CONDITIONAL_ICONS_3_ARROWS_GRAY
            }
            ConditionalIconType::Icons3Flags => {
                libxlsxwriter_sys::lxw_conditional_icon_types_LXW_CONDITIONAL_ICONS_3_FLAGS
            }
            ConditionalIconType::Icons3TrafficLightsUnrimmed => {
                libxlsxwriter_sys::lxw_conditional_icon_types_LXW_CONDITIONAL_ICONS_3_TRAFFIC_LIGHTS_UNRIMMED
            }
            ConditionalIconType::Icons3TrafficLightsRimmed => {
                libxlsxwriter_sys::lxw_conditional_icon_types_LXW_CONDITIONAL_ICONS_3_TRAFFIC_LIGHTS_RIMMED
            }
            ConditionalIconType::Icons3Signs => {
                libxlsxwriter_sys::lxw_conditional_icon_types_LXW_CONDITIONAL_ICONS_3_SIGNS
            }
            ConditionalIconType::Icons3SymbolsCircled => {
                libxlsxwriter_sys::lxw_conditional_icon_types_LXW_CONDITIONAL_ICONS_3_SYMBOLS_CIRCLED
            }
            ConditionalIconType::Icons3SymbolsUncircled => {
                libxlsxwriter_sys::lxw_conditional_icon_types_LXW_CONDITIONAL_ICONS_3_SYMBOLS_UNCIRCLED
            }
            ConditionalIconType::Icons4ArrowColored => {
                libxlsxwriter_sys::lxw_conditional_icon_types_LXW_CONDITIONAL_ICONS_4_ARROWS_COLORED
            }
            ConditionalIconType::Icons4ArrowGray => {
                libxlsxwriter_sys::lxw_conditional_icon_types_LXW_CONDITIONAL_ICONS_4_ARROWS_GRAY
            }
            ConditionalIconType::Icons4RedToBlack => {
                libxlsxwriter_sys::lxw_conditional_icon_types_LXW_CONDITIONAL_ICONS_4_RED_TO_BLACK
            }
            ConditionalIconType::Icons4Rating => {
                libxlsxwriter_sys::lxw_conditional_icon_types_LXW_CONDITIONAL_ICONS_4_RATINGS
            }
            ConditionalIconType::Icons4TrafficLights => {
                libxlsxwriter_sys::lxw_conditional_icon_types_LXW_CONDITIONAL_ICONS_4_TRAFFIC_LIGHTS
            }
            ConditionalIconType::Icons5ArrowColored => {
                libxlsxwriter_sys::lxw_conditional_icon_types_LXW_CONDITIONAL_ICONS_5_ARROWS_COLORED
            }
            ConditionalIconType::Icons5ArrowGray => {
                libxlsxwriter_sys::lxw_conditional_icon_types_LXW_CONDITIONAL_ICONS_5_ARROWS_GRAY
            }
            ConditionalIconType::Icons5Rating => {
                libxlsxwriter_sys::lxw_conditional_icon_types_LXW_CONDITIONAL_ICONS_5_RATINGS
            }
            ConditionalIconType::Icons5Quarters => {
                libxlsxwriter_sys::lxw_conditional_icon_types_LXW_CONDITIONAL_ICONS_5_QUARTERS
            }
        };
        val as u8
    }
}

/// The Icon Set type is used to specify a conditional format with a set of icons such as traffic lights or arrows.
#[derive(Debug, Clone, PartialEq, PartialOrd, Copy)]
pub struct ConditionalIconSet {
    pub style: ConditionalIconType,
    pub reverse_icons: bool,
    pub icons_only: bool,
}

impl ConditionalIconSet {
    pub fn new() -> Self {
        ConditionalIconSet {
            style: ConditionalIconType::Icons5Rating,
            reverse_icons: false,
            icons_only: false,
        }
    }

    pub fn style(mut self, style: ConditionalIconType) -> Self {
        self.style = style;
        self
    }

    pub fn reverse_icons(mut self, reverse_icons: bool) -> Self {
        self.reverse_icons = reverse_icons;
        self
    }

    pub fn icons_only(mut self, icons_only: bool) -> Self {
        self.icons_only = icons_only;
        self
    }

    pub(crate) fn into_internal_value(
        &self,
        conditional_format: &mut libxlsxwriter_sys::lxw_conditional_format,
    ) -> Result<(), XlsxError> {
        conditional_format.type_ =
            libxlsxwriter_sys::lxw_conditional_format_types_LXW_CONDITIONAL_TYPE_ICON_SETS as u8;
        conditional_format.icon_style = self.style.into_internal_type();
        conditional_format.reverse_icons = convert_bool(self.reverse_icons);
        conditional_format.icons_only = convert_bool(self.icons_only);

        Ok(())
    }
}

impl From<ConditionalIconSet> for ConditionalFormat {
    fn from(value: ConditionalIconSet) -> Self {
        ConditionalFormat::IconSet(value)
    }
}

impl ConditionalFormat {
    /// Icon Set
    ///
    /// Example:
    ///
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # use xlsxwriter::worksheet::conditional_format::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-conditional_format-icon-set.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # for i in 0..5 {
    /// #     for j in 0..2 {
    /// #         let v: f64 = i.into();
    /// #         worksheet.write_number(i, j, v, None)?;
    /// #     }
    /// # }
    /// worksheet.conditional_format_range(
    ///     0, 0, 4, 0,
    ///     &ConditionalFormat::icon_set(
    ///         &ConditionalIconSet::new(),
    ///     )
    /// )?;
    /// worksheet.conditional_format_range(
    ///     0, 1, 3, 1,
    ///     &ConditionalIconSet::new().style(ConditionalIconType::Icons4ArrowColored).into()
    /// )?;
    /// # Ok(())
    /// # }
    /// ```    

    pub fn icon_set(icon_set: &ConditionalIconSet) -> ConditionalFormat {
        ConditionalFormat::IconSet(icon_set.clone())
    }
}
