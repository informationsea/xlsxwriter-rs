use super::structs::*;
use crate::{convert_str, Workbook, WorksheetCol, WorksheetRow};

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
    /// # fn main() -> Result<(), XlsxError> {
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
    /// # fn main() -> Result<(), XlsxError> {
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
    /// # fn main() -> Result<(), XlsxError> {
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

    /// The `ChartSeries.set_name_range()` function can be used to set a series name range and is an alternative to using `ChartSeries.set_name()` and a string formula:
    /// ```rust
    /// # use xlsxwriter::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-chart_series-set_name_range-1.xlsx");
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # write_worksheet(&mut worksheet)?; // write worksheet contents
    /// # let mut chart = workbook.add_chart(ChartType::Column);
    /// let mut series = chart.add_series(None, Some("=Sheet1!$B$2:$B$6"));
    /// series.set_name_range("Sheet1", 0, 1); // =Sheet1!$B$1
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

    /// Set the line/border properties of a chart series:
    /// ```rust
    /// # use xlsxwriter::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-chart_series-set_line-1.xlsx");
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # write_worksheet(&mut worksheet)?; // write worksheet contents
    /// # let mut chart = workbook.add_chart(ChartType::Column);
    /// let mut series1 = chart.add_series(None, Some("=Sheet1!$A$2:$A$6"));
    /// let mut series2 = chart.add_series(None, Some("=Sheet1!$B$2:$B$6"));
    /// let mut series3 = chart.add_series(None, Some("=Sheet1!$C$2:$C$6"));
    /// let mut chart_line = ChartLine::new();
    /// chart_line.color = FormatColor::Red;
    /// series1.set_line(&chart_line);
    /// series2.set_line(&chart_line);
    /// series3.set_line(&chart_line);
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
    /// ![Result Image](https://github.com/informationsea/xlsxwriter-rs/raw/master/images/test-chart_series-set_line-1.png)
    pub fn set_line(&mut self, line: &ChartLine) {
        unsafe {
            libxlsxwriter_sys::chart_series_set_line(self.chart_series, &mut line.value());
        }
    }

    /// Set the fill properties of a chart series:
    /// ```rust
    /// # use xlsxwriter::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-chart_series-set_fill-1.xlsx");
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # write_worksheet(&mut worksheet)?; // write worksheet contents
    /// # let mut chart = workbook.add_chart(ChartType::Column);
    /// # let mut series1 = chart.add_series(None, Some("=Sheet1!$A$2:$A$6"));
    /// # let mut series2 = chart.add_series(None, Some("=Sheet1!$B$2:$B$6"));
    /// # let mut series3 = chart.add_series(None, Some("=Sheet1!$C$2:$C$6"));
    /// let mut chart_fill_1 = ChartFill::new();
    /// chart_fill_1.color = FormatColor::Red;
    /// let mut chart_fill_2 = ChartFill::new();
    /// chart_fill_2.color = FormatColor::Yellow;
    /// let mut chart_fill_3 = ChartFill::new();
    /// chart_fill_3.color = FormatColor::Green;
    /// series1.set_fill(&chart_fill_1);
    /// series2.set_fill(&chart_fill_2);
    /// series3.set_fill(&chart_fill_3);
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
    /// ![Result Image](https://github.com/informationsea/xlsxwriter-rs/raw/master/images/test-chart_series-set_fill-1.png)
    pub fn set_fill(&mut self, fill: &ChartFill) {
        unsafe {
            libxlsxwriter_sys::chart_series_set_fill(self.chart_series, &mut fill.value());
        }
    }

    /// Invert the fill color for negative values. Usually only applicable to column and bar charts.
    /// ```rust
    /// # use xlsxwriter::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-chart_series-set_invert_if_negative-1.xlsx");
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # write_worksheet(&mut worksheet)?; // write worksheet contents
    /// # let mut chart = workbook.add_chart(ChartType::Column);
    /// # let mut series1 = chart.add_series(None, Some("=Sheet1!$A$2:$A$6"));
    /// # let mut series2 = chart.add_series(None, Some("=Sheet1!$B$2:$B$6"));
    /// # let mut series3 = chart.add_series(None, Some("=Sheet1!$C$2:$C$6"));
    /// # series1.set_name("=Sheet1!$A$1");
    /// # series2.set_name("=Sheet1!$B$1");
    /// # series3.set_name("=Sheet1!$C$1");
    /// # let mut chart_fill_1 = ChartFill::new();
    /// # chart_fill_1.color = FormatColor::Red;
    /// # let mut chart_fill_2 = ChartFill::new();
    /// # chart_fill_2.color = FormatColor::Yellow;
    /// # let mut chart_fill_3 = ChartFill::new();
    /// # chart_fill_3.color = FormatColor::Green;
    /// # series1.set_fill(&chart_fill_1);
    /// series1.set_invert_if_negative();
    /// # series2.set_fill(&chart_fill_2);
    /// # series2.set_invert_if_negative();
    /// # series3.set_fill(&chart_fill_3);
    /// # series3.set_invert_if_negative();
    /// # worksheet.insert_chart(1, 3, &chart)?;
    /// # workbook.close()
    /// # }
    /// # fn write_worksheet(worksheet: &mut Worksheet) -> Result<(), XlsxError> {
    /// # worksheet.write_string(0, 0, "Set 1", None)?;
    /// # worksheet.write_string(0, 1, "Set 2", None)?;
    /// # worksheet.write_string(0, 2, "Set 3", None)?;
    /// # for i in 1..6 {
    /// #     let j: f64 = i.into();
    /// #     worksheet.write_number(i, 0, (j*10.) - 20., None)?;
    /// #     worksheet.write_number(i, 1, (j*10. + 2.) - 20., None)?;
    /// #     worksheet.write_number(i, 2, (j*10. + 4.) - 20., None)?;
    /// # }
    /// # Ok(())
    /// # }
    /// ```
    pub fn set_invert_if_negative(&mut self) {
        unsafe {
            libxlsxwriter_sys::chart_series_set_invert_if_negative(self.chart_series);
        }
    }

    /// Set the pattern properties of a chart series:
    /// ```rust
    /// # use xlsxwriter::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-chart_series-set_pattern-1.xlsx");
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # write_worksheet(&mut worksheet)?; // write worksheet contents
    /// # let mut chart = workbook.add_chart(ChartType::Column);
    /// # let mut series1 = chart.add_series(None, Some("=Sheet1!$A$2:$A$6"));
    /// # let mut series2 = chart.add_series(None, Some("=Sheet1!$B$2:$B$6"));
    /// # series1.set_name("=Sheet1!$A$1");
    /// # series2.set_name("=Sheet1!$B$1");
    /// let pattern1 = ChartPattern::new(FormatColor::Custom(0x804000), FormatColor::Custom(0xC68C53), ChartPatternType::Shingle);
    /// series1.set_pattern(&pattern1);
    /// let pattern2 = ChartPattern::new(FormatColor::Custom(0xB30000), FormatColor::Custom(0xFF6666), ChartPatternType::HorizontalBrick);
    /// series2.set_pattern(&pattern2);
    /// # worksheet.insert_chart(1, 3, &chart)?;
    /// # workbook.close()
    /// # }
    /// # fn write_worksheet(worksheet: &mut Worksheet) -> Result<(), XlsxError> {
    /// # worksheet.write_string(0, 0, "Shingle", None)?;
    /// # worksheet.write_string(0, 1, "Brick", None)?;
    /// # for i in 1..6 {
    /// #     let j: f64 = i.into();
    /// #     worksheet.write_number(i, 0, (j*10.) - 20., None)?;
    /// #     worksheet.write_number(i, 1, (j*10. + 2.) - 20., None)?;
    /// # }
    /// # Ok(())
    /// # }
    /// ```    
    pub fn set_pattern(&mut self, pattern: &ChartPattern) {
        unsafe {
            libxlsxwriter_sys::chart_series_set_pattern(self.chart_series, &mut pattern.value())
        }
    }
}
