#[derive(Debug, Copy, Clone, PartialEq, PartialOrd)]
pub enum ChartType {
    None,
    Area,
    AreaStacked,
    AreaStackedPercent,
    Bar,
    BarStacked,
    Column,
    ColumnStacked,
    ColumnStackedPercent,
    Doughnut,
    Line,
    Pie,
    Scatter,
    ScatterStraight,
    ScatterStraightWithMarkers,
    ScatterSmooth,
    ScatterSmoothWithMarkers,
    Radar,
    RadarWithMarkers,
    RadarFilled,
}

impl ChartType {
    pub(crate) fn value(self) -> u8 {
        let value = match self {
            ChartType::None => libxlsxwriter_sys::lxw_chart_type_LXW_CHART_NONE,
            ChartType::Area => libxlsxwriter_sys::lxw_chart_type_LXW_CHART_AREA,
            ChartType::AreaStacked => libxlsxwriter_sys::lxw_chart_type_LXW_CHART_AREA_STACKED,
            ChartType::AreaStackedPercent => {
                libxlsxwriter_sys::lxw_chart_type_LXW_CHART_AREA_STACKED_PERCENT
            }
            ChartType::Bar => libxlsxwriter_sys::lxw_chart_type_LXW_CHART_BAR,
            ChartType::BarStacked => libxlsxwriter_sys::lxw_chart_type_LXW_CHART_BAR_STACKED,
            ChartType::Column => libxlsxwriter_sys::lxw_chart_type_LXW_CHART_COLUMN,
            ChartType::ColumnStacked => libxlsxwriter_sys::lxw_chart_type_LXW_CHART_COLUMN_STACKED,
            ChartType::ColumnStackedPercent => {
                libxlsxwriter_sys::lxw_chart_type_LXW_CHART_COLUMN_STACKED_PERCENT
            }
            ChartType::Doughnut => libxlsxwriter_sys::lxw_chart_type_LXW_CHART_DOUGHNUT,
            ChartType::Line => libxlsxwriter_sys::lxw_chart_type_LXW_CHART_LINE,
            ChartType::Pie => libxlsxwriter_sys::lxw_chart_type_LXW_CHART_PIE,
            ChartType::Scatter => libxlsxwriter_sys::lxw_chart_type_LXW_CHART_SCATTER,
            ChartType::ScatterStraight => {
                libxlsxwriter_sys::lxw_chart_type_LXW_CHART_SCATTER_STRAIGHT
            }
            ChartType::ScatterStraightWithMarkers => {
                libxlsxwriter_sys::lxw_chart_type_LXW_CHART_SCATTER_STRAIGHT_WITH_MARKERS
            }
            ChartType::ScatterSmooth => libxlsxwriter_sys::lxw_chart_type_LXW_CHART_SCATTER_SMOOTH,
            ChartType::ScatterSmoothWithMarkers => {
                libxlsxwriter_sys::lxw_chart_type_LXW_CHART_SCATTER_SMOOTH_WITH_MARKERS
            }
            ChartType::Radar => libxlsxwriter_sys::lxw_chart_type_LXW_CHART_RADAR,
            ChartType::RadarWithMarkers => {
                libxlsxwriter_sys::lxw_chart_type_LXW_CHART_RADAR_WITH_MARKERS
            }
            ChartType::RadarFilled => libxlsxwriter_sys::lxw_chart_type_LXW_CHART_RADAR_FILLED,
        };
        value as u8
    }
}

#[derive(Copy, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum ChartDashType {
    Solid,
    RoundDot,
    SquareDot,
    Dash,
    DashDot,
    LongDash,
    LongDashDot,
    LongDashDotDot,
}

impl ChartDashType {
    pub(crate) fn value(self) -> u8 {
        let value = match self {
            ChartDashType::Solid => {
                libxlsxwriter_sys::lxw_chart_line_dash_type_LXW_CHART_LINE_DASH_SOLID
            }
            ChartDashType::RoundDot => {
                libxlsxwriter_sys::lxw_chart_line_dash_type_LXW_CHART_LINE_DASH_ROUND_DOT
            }
            ChartDashType::SquareDot => {
                libxlsxwriter_sys::lxw_chart_line_dash_type_LXW_CHART_LINE_DASH_SQUARE_DOT
            }
            ChartDashType::Dash => {
                libxlsxwriter_sys::lxw_chart_line_dash_type_LXW_CHART_LINE_DASH_DASH
            }
            ChartDashType::DashDot => {
                libxlsxwriter_sys::lxw_chart_line_dash_type_LXW_CHART_LINE_DASH_DASH_DOT
            }
            ChartDashType::LongDash => {
                libxlsxwriter_sys::lxw_chart_line_dash_type_LXW_CHART_LINE_DASH_LONG_DASH
            }
            ChartDashType::LongDashDot => {
                libxlsxwriter_sys::lxw_chart_line_dash_type_LXW_CHART_LINE_DASH_LONG_DASH_DOT
            }
            ChartDashType::LongDashDotDot => {
                libxlsxwriter_sys::lxw_chart_line_dash_type_LXW_CHART_LINE_DASH_LONG_DASH_DOT_DOT
            }
        };
        value as u8
    }
}

#[derive(Copy, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum ChartPatternType {
    /// None pattern.
    None,
    /// 5 Percent pattern.
    Percent5,
    /// 10 Percent pattern.
    Percent10,
    /// 20 Percent pattern.
    Percent20,
    /// 25 Percent pattern.
    Percent25,
    /// 30 Percent pattern.
    Percent30,
    /// 40 Percent pattern.
    Percent40,
    /// 50 Percent pattern.
    Percent50,
    /// 60 Percent pattern.
    Percent60,
    /// 70 Percent pattern.
    Percent70,
    /// 75 Percent pattern.
    Percent75,
    /// 80 Percent pattern.
    Percent80,
    /// 90 Percent pattern.
    Percent90,
    /// Light downward diagonal pattern.
    LightDownwardDiagonal,
    /// Light upward diagonal pattern.
    LightUpwardDiagonal,
    /// Dark downward diagonal pattern.
    DarkDownwardDiagonal,
    /// Dark upward diagonal pattern.
    DarkUpwardDiagonal,
    /// Wide downward diagonal pattern.
    WideDownwardDiagonal,
    /// Wide upward diagonal pattern.
    WideUpwardDiagonal,
    /// Light vertical pattern.
    LightVertical,
    /// Light horizontal pattern.
    LightHorizontal,
    /// Narrow vertical pattern.
    NarrowVertical,
    /// Narrow horizontal pattern.
    NarrowHorizontal,
    /// Dark vertical pattern.
    DarkVertical,
    /// Dark horizontal pattern.
    DarkHorizontal,
    /// Dashed downward diagonal pattern.
    DashedDownwardDiagonal,
    /// Dashed upward diagonal pattern.
    DashedUpwardDiagonal,
    /// Dashed horizontal pattern.
    DashedHorizontal,
    /// Dashed vertical pattern.
    DashedVertical,
    /// Small confetti pattern.
    SmallConfetti,
    /// Large confetti pattern.
    LargeConfetti,
    /// Zigzag pattern.    
    Zigzag,
    /// Wave pattern.
    Wave,
    /// Diagonal brick pattern.
    DiagonalBrick,
    ///Horizontal brick pattern.
    HorizontalBrick,
    /// Weave pattern.
    Weave,
    /// Plaid pattern.
    Plaid,
    /// Divot pattern.
    Divot,
    /// Dotted grid pattern.
    DottedGrid,
    /// Dotted diamond pattern.
    DottedDiamond,
    /// Shingle pattern.
    Shingle,
    /// Trellis pattern.
    Trellis,
    /// Sphere pattern.
    Sphere,
    /// Small grid pattern.
    SmallGrid,
    /// Large grid pattern.
    LargeGrid,
    /// Small check pattern.
    SmallCheck,
    /// Large check pattern.
    LargeCheck,
    /// Outlined diamond pattern.
    OutlinedDiamond,
    /// Solid diamond pattern.
    SolidDiamond,
}

impl ChartPatternType {
    pub(crate) fn value(self) -> u8 {
        let value = match self {
            ChartPatternType::None => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_NONE
            }
            ChartPatternType::Percent5 => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_PERCENT_5
            }
            ChartPatternType::Percent10 => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_PERCENT_10
            }
            ChartPatternType::Percent20 => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_PERCENT_20
            }
            ChartPatternType::Percent25 => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_PERCENT_25
            }
            ChartPatternType::Percent30 => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_PERCENT_30
            }
            ChartPatternType::Percent40 => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_PERCENT_40
            }
            ChartPatternType::Percent50 => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_PERCENT_50
            }
            ChartPatternType::Percent60 => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_PERCENT_60
            }
            ChartPatternType::Percent70 => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_PERCENT_70
            }
            ChartPatternType::Percent75 => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_PERCENT_75
            }
            ChartPatternType::Percent80 => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_PERCENT_80
            }
            ChartPatternType::Percent90 => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_PERCENT_90
            }
            ChartPatternType::LightDownwardDiagonal => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_LIGHT_DOWNWARD_DIAGONAL
            }
            ChartPatternType::LightUpwardDiagonal => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_LIGHT_UPWARD_DIAGONAL
            }
            ChartPatternType::DarkDownwardDiagonal => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_DARK_DOWNWARD_DIAGONAL
            }
            ChartPatternType::DarkUpwardDiagonal => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_DARK_UPWARD_DIAGONAL
            }
            ChartPatternType::WideDownwardDiagonal => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_WIDE_DOWNWARD_DIAGONAL
            }
            ChartPatternType::WideUpwardDiagonal => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_WIDE_UPWARD_DIAGONAL
            }
            ChartPatternType::LightVertical => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_LIGHT_VERTICAL
            }
            ChartPatternType::LightHorizontal => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_LIGHT_HORIZONTAL
            }
            ChartPatternType::NarrowVertical => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_NARROW_VERTICAL
            }
            ChartPatternType::NarrowHorizontal => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_NARROW_HORIZONTAL
            }
            ChartPatternType::DarkVertical => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_DARK_VERTICAL
            }
            ChartPatternType::DarkHorizontal => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_DARK_HORIZONTAL
            }
            ChartPatternType::DashedDownwardDiagonal => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_DASHED_DOWNWARD_DIAGONAL
            }
            ChartPatternType::DashedUpwardDiagonal => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_DASHED_UPWARD_DIAGONAL
            }
            ChartPatternType::DashedHorizontal => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_DASHED_HORIZONTAL
            }
            ChartPatternType::DashedVertical => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_DASHED_VERTICAL
            }
            ChartPatternType::SmallConfetti => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_SMALL_CONFETTI
            }
            ChartPatternType::LargeConfetti => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_LARGE_CONFETTI
            }
            ChartPatternType::Zigzag => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_ZIGZAG
            }
            ChartPatternType::Wave => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_WAVE
            }
            ChartPatternType::DiagonalBrick => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_DIAGONAL_BRICK
            }
            ChartPatternType::HorizontalBrick => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_HORIZONTAL_BRICK
            }
            ChartPatternType::Weave => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_WEAVE
            }
            ChartPatternType::Plaid => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_PLAID
            }
            ChartPatternType::Divot => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_DIVOT
            }
            ChartPatternType::DottedGrid => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_DOTTED_GRID
            }
            ChartPatternType::DottedDiamond => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_DOTTED_DIAMOND
            }
            ChartPatternType::Shingle => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_SHINGLE
            }
            ChartPatternType::Trellis => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_TRELLIS
            }
            ChartPatternType::Sphere => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_SPHERE
            }
            ChartPatternType::SmallGrid => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_SMALL_GRID
            }
            ChartPatternType::LargeGrid => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_LARGE_GRID
            }
            ChartPatternType::SmallCheck => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_SMALL_CHECK
            }
            ChartPatternType::LargeCheck => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_LARGE_CHECK
            }
            ChartPatternType::OutlinedDiamond => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_OUTLINED_DIAMOND
            }
            ChartPatternType::SolidDiamond => {
                libxlsxwriter_sys::lxw_chart_pattern_type_LXW_CHART_PATTERN_SOLID_DIAMOND
            }
        };
        value as u8
    }
}
