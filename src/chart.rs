use super::{convert_bool, convert_str, FormatColor, Workbook, WorksheetCol, WorksheetRow};

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
    fn value(self) -> u8 {
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

    fn value(&self) -> libxlsxwriter_sys::lxw_chart_line {
        libxlsxwriter_sys::lxw_chart_line {
            color: self.color.value(),
            none: convert_bool(self.none),
            width: self.width,
            dash_type: self.dash_type.value(),
            transparency: self.transparency,
            has_color: convert_bool(false),
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

/// Struct to represent an Excel chart data series.
/// This struct is created using the chart.add_series() function. It is used in functions that modify a chart series but the members of the struct aren't modified directly.
pub struct ChartSeries<'a> {
    pub(crate) _workbook: &'a Workbook,
    pub(crate) chart_series: *mut libxlsxwriter_sys::lxw_chart_series,
}

impl<'a> ChartSeries<'a> {
    /// The categories and values of a chart data series are generally set using the chart_add_series() function and Excel range formulas like "=Sheet1!$A$2:$A$7".
    ///
    /// The `ChartSeries.set_categories()` function is an alternative method that is easier to generate programmatically. It requires that you set the categories and values parameters in Chart.add_series() to `None` and then set them using row and column values in ChartSeries.set_categories() and ChartSeries.set_values():
    /// ```rust
    /// # use xlsxwriter::*;
    /// # fn main() { let _ = run(); }
    /// # fn run() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-chart_series-set_categories-1.xlsx");
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # write_worksheet(&mut worksheet)?; // write worksheet contents
    /// # let mut chart = workbook.add_chart(ChartType::Column);
    /// let mut series = chart.add_series(None, None);
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
    pub fn set_categories(
        &mut self,
        sheet_name: &str,
        first_row: WorksheetRow,
        first_column: WorksheetCol,
        last_row: WorksheetRow,
        last_column: WorksheetCol,
    ) {
        let sheet_name_vec = convert_str(sheet_name);
        unsafe {
            libxlsxwriter_sys::chart_series_set_categories(
                self.chart_series,
                sheet_name_vec.as_ptr() as *const i8,
                first_row,
                first_column,
                last_row,
                last_column,
            );
        }
        self._workbook.const_str.borrow_mut().push(sheet_name_vec);
    }

    /// The categories and values of a chart data series are generally set using the `Chart.add_series()` function and Excel range formulas like "=Sheet1!$A$2:$A$7".
    ///
    /// The `Chart.series_set_values()` function is an alternative method that is easier to generate programmatically. See the documentation for `ChartSeries.set_categories()` above.
    pub fn set_values(
        &mut self,
        sheet_name: &str,
        first_row: WorksheetRow,
        first_column: WorksheetCol,
        last_row: WorksheetRow,
        last_column: WorksheetCol,
    ) {
        let sheet_name_vec = convert_str(sheet_name);
        unsafe {
            libxlsxwriter_sys::chart_series_set_values(
                self.chart_series,
                sheet_name_vec.as_ptr() as *const i8,
                first_row,
                first_column,
                last_row,
                last_column,
            );
        }
        self._workbook.const_str.borrow_mut().push(sheet_name_vec);
    }

    /// This function is used to set the name for a chart data series. The series name in Excel is displayed in the chart legend and in the formula bar. The name property is optional and if it isn't supplied it will default to `Series 1..n`.
    ///
    /// ```rust
    /// use xlsxwriter::*;
    /// # fn main() { let _ = run(); }
    /// # fn run() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-chart_series-set_name-1.xlsx");
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # write_worksheet(&mut worksheet)?; // write worksheet contents
    /// # let mut chart = workbook.add_chart(ChartType::Column);
    /// let mut series = chart.add_series(None, Some("=Sheet1!$A$2:$A$6"));
    /// series.set_name("Quarterly budget data");
    /// # worksheet.insert_chart(1, 3, &chart)?;
    /// # workbook.close()
    /// # }
    /// # fn write_worksheet(worksheet: &mut Worksheet) -> Result<(), XlsxError> {
    /// # worksheet.write_string(0, 0, "Set 1", None)?;
    /// # worksheet.write_string(0, 1, "Set 2", None)?;
    /// # worksheet.write_string(0, 2, "Set 3", None)?;
    /// # for i in 1..6 {
    /// #     worksheet.write_number(i, 0, (i*10).into(), None)?;
    /// #     worksheet.write_number(i, 1, (i*10 + 2).into(), None)?;
    /// #     worksheet.write_number(i, 2, (i*10 + 4).into(), None)?;
    /// # }
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// The name parameter can also be a formula such as =Sheet1!$A$1 to point to a cell in the workbook that contains the name:
    /// ```rust
    /// use xlsxwriter::*;
    /// # fn main() { let _ = run(); }
    /// # fn run() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-chart_series-set_name-2.xlsx");
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # write_worksheet(&mut worksheet)?; // write worksheet contents
    /// # let mut chart = workbook.add_chart(ChartType::Column);
    /// let mut series = chart.add_series(None, Some("=Sheet1!$A$2:$A$6"));
    /// series.set_name("=Sheet1!$A$1:$A$1");
    /// # worksheet.insert_chart(1, 3, &chart)?;
    /// # workbook.close()
    /// # }
    /// # fn write_worksheet(worksheet: &mut Worksheet) -> Result<(), XlsxError> {
    /// # worksheet.write_string(0, 0, "Set 1", None)?;
    /// # worksheet.write_string(0, 1, "Set 2", None)?;
    /// # worksheet.write_string(0, 2, "Set 3", None)?;
    /// # for i in 1..6 {
    /// #     worksheet.write_number(i, 0, (i*10).into(), None)?;
    /// #     worksheet.write_number(i, 1, (i*10 + 2).into(), None)?;
    /// #     worksheet.write_number(i, 2, (i*10 + 4).into(), None)?;
    /// # }
    /// # Ok(())
    /// # }
    /// ```
    pub fn set_name(&mut self, name: &str) {
        let name_vec = convert_str(name);
        unsafe {
            libxlsxwriter_sys::chart_series_set_name(
                self.chart_series,
                name_vec.as_ptr() as *const i8,
            );
        }
        self._workbook.const_str.borrow_mut().push(name_vec);
    }

    pub fn set_name_range(&mut self, sheet_name: &str, row: WorksheetRow, column: WorksheetCol) {
        let sheet_name_vec = convert_str(sheet_name);
        unsafe {
            libxlsxwriter_sys::chart_series_set_name_range(
                self.chart_series,
                sheet_name_vec.as_ptr() as *const i8,
                row,
                column,
            );
        }
        self._workbook.const_str.borrow_mut().push(sheet_name_vec);
    }

    pub fn set_line(&mut self, line: &ChartLine) {
        unsafe {
            libxlsxwriter_sys::chart_series_set_line(self.chart_series, &mut line.value());
        }
    }
}

/// The Chart object represents an Excel chart. It provides functions for adding data series to the chart and for configuring the chart.
///
/// A Chart object isn't created directly. Instead a chart is created by calling the Workbook.add_chart() function from a Workbook object. For example:
/// ```rust
/// use xlsxwriter::*;
/// # fn main() { let _ = run(); }
/// # fn run() -> Result<(), XlsxError> {
/// let workbook = Workbook::new("test-chart.xlsx");
/// let mut worksheet = workbook.add_worksheet(None)?;
/// write_worksheet(&mut worksheet)?; // write worksheet contents
/// let mut chart = workbook.add_chart(ChartType::Column);
/// chart.add_series(None, Some("=Sheet1!$A$1:$A$5"));
/// chart.add_series(None, Some("=Sheet1!$B$1:$B$5"));
/// chart.add_series(None, Some("=Sheet1!$C$1:$C$5"));
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
    /// # use xlsxwriter::*;
    /// # fn main() { let _ = run(); }
    /// # fn run() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-chart-add_series-1.xlsx");
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # write_worksheet(&mut worksheet)?; // write worksheet contents
    /// # let mut chart = workbook.add_chart(ChartType::Column);
    /// chart.add_series(Some("=Sheet1!$A$1:$A$5"), Some("=Sheet1!$B$1:$B$5"));
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
    /// # use xlsxwriter::*;
    /// # fn main() { let _ = run(); }
    /// # fn run() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-chart-add_series-2.xlsx");
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # write_worksheet(&mut worksheet)?; // write worksheet contents
    /// # let mut chart = workbook.add_chart(ChartType::Column);
    /// chart.add_series(None, Some("=Sheet1!$B$1:$B$5"));
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
    /// # use xlsxwriter::*;
    /// # fn main() { let _ = run(); }
    /// # fn run() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-chart-add_series-3.xlsx");
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # write_worksheet(&mut worksheet)?; // write worksheet contents
    /// # let mut chart = workbook.add_chart(ChartType::Column);
    /// let mut series = chart.add_series(None, None);
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
    /// # use xlsxwriter::*;
    /// # fn main() { let _ = run(); }
    /// # fn run() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-chart-add_series-4.xlsx");
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
    /// # use xlsxwriter::*;
    /// # fn main() { let _ = run(); }
    /// # fn run() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-chart-add_series-5.xlsx");
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # write_worksheet(&mut worksheet)?; // write worksheet contents
    /// # let mut chart = workbook.add_chart(ChartType::Column);
    /// chart.add_series(Some("=(Sheet1!$A$1:$A$5,Sheet1!$A$10:$A$18)"), Some("=(Sheet1!$B$1:$B$5,Sheet1!$B$10:$B$18)"));
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
    ) -> ChartSeries<'a> {
        let categories_vec = categories.map(convert_str);
        let values_vec = values.map(convert_str);
        let mut const_str = self._workbook.const_str.borrow_mut();
        let series = unsafe {
            libxlsxwriter_sys::chart_add_series(
                self.chart,
                categories_vec
                    .as_ref()
                    .map(|x| x.as_ptr())
                    .unwrap_or(std::ptr::null()) as *const i8,
                values_vec
                    .as_ref()
                    .map(|x| x.as_ptr())
                    .unwrap_or(std::ptr::null()) as *const i8,
            )
        };
        if let Some(x) = categories_vec {
            const_str.push(x);
        }
        if let Some(x) = values_vec {
            const_str.push(x);
        }
        ChartSeries {
            _workbook: self._workbook,
            chart_series: series,
        }
    }
}
