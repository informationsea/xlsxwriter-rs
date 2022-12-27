mod constants;
mod series;
mod structs;

use crate::format::FormatColor;
use crate::XlsxError;

pub use self::constants::*;
pub use self::series::*;
pub use self::structs::*;
use super::Workbook;

/// The Chart object represents an Excel chart. It provides functions for adding data series to the chart and for configuring the chart.
///
/// A Chart object isn't created directly. Instead a chart is created by calling the Workbook.add_chart() function from a Workbook object. For example:
/// ```rust
/// use xlsxwriter::prelude::*;
/// # fn main() -> Result<(), XlsxError> {
/// let workbook = Workbook::new("test-chart.xlsx")?;
/// let mut worksheet = workbook.add_worksheet(None)?;
/// write_worksheet(&mut worksheet)?; // write worksheet contents
/// let mut chart = workbook.add_chart(ChartType::Column);
/// chart.add_series(None, Some("=Sheet1!$A$1:$A$5"))?;
/// chart.add_series(None, Some("=Sheet1!$B$1:$B$5"))?;
/// chart.add_series(None, Some("=Sheet1!$C$1:$C$5"))?;
/// worksheet.insert_chart(1, 3, &chart)?;
/// workbook.close()
/// # }
/// # fn write_worksheet(worksheet: &mut Worksheet) -> Result<(), XlsxError> {
/// # for i in 0..5 {
/// #     worksheet.write_number(i, 0, (i*10).into(), None)?;
/// #     worksheet.write_number(i, 1, (i*10 + 2).into(), None)?;
/// #     worksheet.write_number(i, 2, (i*10 + 4).into(), None)?;
/// # }
/// # Ok(())
/// # }
/// ```
/// The chart in the worksheet will look like this:
/// ![Result Image](https://github.com/informationsea/xlsxwriter-rs/raw/master/images/test-chart-1.png)
///
///
/// The basic procedure for adding a chart to a worksheet is:
///
/// Create the chart with Workbook.add_chart().
/// Add one or more data series to the chart which refers to data in the workbook using Chart.add_series().
/// Configure the chart with the other available functions shown below.
/// Insert the chart into a worksheet using Worksheet.insert_chart().
pub struct Chart<'a> {
    pub(crate) _workbook: &'a Workbook,
    pub(crate) chart: *mut libxlsxwriter_sys::lxw_chart,
}

impl<'a> Chart<'a> {
    /// In Excel a chart **series** is a collection of information that defines which data is plotted such as the categories and values. It is also used to define the formatting for the data.
    ///
    /// For an libxlsxwriter chart object the chart_add_series() function is used to set the categories and values of the series:
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-chart-add_series-1.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # write_worksheet(&mut worksheet)?; // write worksheet contents
    /// # let mut chart = workbook.add_chart(ChartType::Column);
    /// chart.add_series(Some("=Sheet1!$A$1:$A$5"), Some("=Sheet1!$B$1:$B$5"))?;
    /// # worksheet.insert_chart(1, 3, &chart)?;
    /// # workbook.close()
    /// # }
    /// # fn write_worksheet(worksheet: &mut Worksheet) -> Result<(), XlsxError> {
    /// # for i in 0..5 {
    /// #     worksheet.write_string(i, 0, &format!("value {}", i + 1), None)?;
    /// #     worksheet.write_number(i, 1, (i*10 + 2).into(), None)?;
    /// # }
    /// # Ok(())
    /// # }
    /// ```
    /// The series parameters are:
    ///
    /// *categories: This sets the chart category labels. The category is more or less the same as the X axis. In most Excel chart types the categories property is optional and the chart will just assume a sequential series from 1..n:
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-chart-add_series-2.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # write_worksheet(&mut worksheet)?; // write worksheet contents
    /// # let mut chart = workbook.add_chart(ChartType::Column);
    /// chart.add_series(None, Some("=Sheet1!$B$1:$B$5"))?;
    /// # worksheet.insert_chart(1, 3, &chart)?;
    /// # workbook.close()
    /// # }
    /// # fn write_worksheet(worksheet: &mut Worksheet) -> Result<(), XlsxError> {
    /// # for i in 0..5 {
    /// #     worksheet.write_string(i, 0, &format!("value {}", i + 1), None)?;
    /// #     worksheet.write_number(i, 1, (i*10 + 2).into(), None)?;
    /// # }
    /// # Ok(())
    /// # }
    /// ```
    /// * values: This is the most important property of a series and is the only mandatory option for every chart object. This parameter links the chart with the worksheet data that it displays.
    ///
    /// The categories and values should be a string formula like "=Sheet1!$A$2:$A$7" in the same way it is represented in Excel. This is convenient when recreating a chart from an example in Excel but it is trickier to generate programmatically. For these cases you can set the categories and values to None and use the ChartSeries.set_categories() and ChartSeries.set_values() functions:
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-chart-add_series-3.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # write_worksheet(&mut worksheet)?; // write worksheet contents
    /// # let mut chart = workbook.add_chart(ChartType::Column);
    /// let mut series = chart.add_series(None, None)?;
    /// series.set_categories("Sheet1", 0, 0, 4, 0); // "=Sheet1!$A$1:$A$5"
    /// series.set_values("Sheet1", 0, 1, 4, 1);     // "=Sheet1!$B$1:$B$5"
    /// # worksheet.insert_chart(1, 3, &chart)?;
    /// # workbook.close()
    /// # }
    /// # fn write_worksheet(worksheet: &mut Worksheet) -> Result<(), XlsxError> {
    /// # for i in 0..5 {
    /// #     worksheet.write_string(i, 0, &format!("value {}", i + 1), None)?;
    /// #     worksheet.write_number(i, 1, (i*10 + 2).into(), None)?;
    /// # }
    /// # Ok(())
    /// # }
    /// ```
    /// As shown in the previous example the return value from Chart.add_series() is a `ChartSeries` struct. This can be used in other functions that configure a series.
    ///
    /// More than one series can be added to a chart. The series numbering and order in the Excel chart will be the same as the order in which they are added in libxlsxwriter:
    ///
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-chart-add_series-4.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # write_worksheet(&mut worksheet)?; // write worksheet contents
    /// # let mut chart = workbook.add_chart(ChartType::Column);
    /// chart.add_series(None, Some("=Sheet1!$A$1:$A$5"));
    /// chart.add_series(None, Some("=Sheet1!$B$1:$B$5"));
    /// chart.add_series(None, Some("=Sheet1!$C$1:$C$5"));
    /// # worksheet.insert_chart(1, 3, &chart)?;
    /// # workbook.close()
    /// # }
    /// # fn write_worksheet(worksheet: &mut Worksheet) -> Result<(), XlsxError> {
    /// # for i in 0..5 {
    /// #     worksheet.write_number(i, 0, (i*10 + 1).into(), None)?;
    /// #     worksheet.write_number(i, 1, (i*10 + 2).into(), None)?;
    /// #     worksheet.write_number(i, 2, (i*10 + 5).into(), None)?;
    /// # }
    /// # Ok(())
    /// # }
    /// ```
    /// It is also possible to specify non-contiguous ranges:
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-chart-add_series-5.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # write_worksheet(&mut worksheet)?; // write worksheet contents
    /// # let mut chart = workbook.add_chart(ChartType::Column);
    /// chart.add_series(Some("=(Sheet1!$A$1:$A$5,Sheet1!$A$10:$A$18)"), Some("=(Sheet1!$B$1:$B$5,Sheet1!$B$10:$B$18)"))?;
    /// # worksheet.insert_chart(1, 3, &chart)?;
    /// # workbook.close()
    /// # }
    /// # fn write_worksheet(worksheet: &mut Worksheet) -> Result<(), XlsxError> {
    /// # for i in 0..20 {
    /// #     worksheet.write_string(i, 0, &format!("value {}", i + 1), None)?;
    /// #     worksheet.write_number(i, 1, (i*10 + 2).into(), None)?;
    /// # }
    /// # Ok(())
    /// # }
    /// ```

    pub fn add_series(
        &mut self,
        categories: Option<&str>,
        values: Option<&str>,
    ) -> Result<ChartSeries<'a>, XlsxError> {
        let series = unsafe {
            libxlsxwriter_sys::chart_add_series(
                self.chart,
                self._workbook.register_option_str(categories)?,
                self._workbook.register_option_str(values)?,
            )
        };
        Ok(ChartSeries {
            _workbook: self._workbook,
            chart_series: series,
        })
    }

    /// The chart_title_set_name() function sets the name (title) for the chart. The name is displayed above the chart.
    /// The name parameter can also be a formula such as `=Sheet1!$A$1` to point to a cell in the workbook that contains the name.
    /// The Excel default is to have no chart title.
    pub fn add_title(&mut self, title: &str) -> Result<(), XlsxError> {
        unsafe {
            libxlsxwriter_sys::chart_title_set_name(self.chart, self._workbook.register_str(title)?)
        }
        Ok(())
    }
}

/// Struct to represent an Excel chart data series.
/// This struct is created using the chart.add_series() function. It is used in functions that modify a chart series but the members of the struct aren't modified directly.
pub struct ChartSeries<'a> {
    pub(crate) _workbook: &'a Workbook,
    pub(crate) chart_series: *mut libxlsxwriter_sys::lxw_chart_series,
}

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

#[derive(Copy, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum ChartMarkerType {
    MarkerAutomatic,
    MarkerNone,
    MarkerSquare,
    MarkerDiamond,
    MarkerTriangle,
    MarkerX,
    MarkerStar,
    MarkerShortDash,
    MarkerLongDash,
    MarkerCircle,
    MarkerPlus,
}
