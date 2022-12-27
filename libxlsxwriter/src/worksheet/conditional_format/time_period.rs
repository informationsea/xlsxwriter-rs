use super::{ConditionalFormat, ConditionalFormatTypes};
use crate::{Format, XlsxError};

/// The Time Period type is used to specify Excel's "Dates Occurring" style conditional format.
///
/// See [`ConditionalFormat::time_period`] to learn more
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub enum ConditionalFormatTimePeriodCriteria {
    /// Format cells with a date of yesterday.    
    Yesterday,
    /// Format cells with a date of today.
    Today,
    /// Format cells with a date of tomorrow.
    Tomorrow,
    /// Format cells with a date in the last 7 days.
    Last7Days,
    /// Format cells with a date in the last week.
    LastWeek,
    /// Format cells with a date in the current week.
    ThisWeek,
    /// Format cells with a date in the next week.
    NextWeek,
    /// Format cells with a date in the last month.
    LastMonth,
    /// Format cells with a date in the current month.
    ThisMonth,
    /// Format cells with a date in the next month.
    NextMonth,
}

impl ConditionalFormatTimePeriodCriteria {
    pub(crate) fn into_internal_value(
        &self,
        conditional_format: &mut libxlsxwriter_sys::lxw_conditional_format,
    ) -> Result<(), XlsxError> {
        conditional_format.type_ =
            libxlsxwriter_sys::lxw_conditional_format_types_LXW_CONDITIONAL_TYPE_TIME_PERIOD as u8;
        match self {
            ConditionalFormatTimePeriodCriteria::Yesterday => {
                conditional_format.criteria = libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_YESTERDAY as u8;
            }
            ConditionalFormatTimePeriodCriteria::Today => {
                conditional_format.criteria = libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_TODAY as u8;
            }
            ConditionalFormatTimePeriodCriteria::Tomorrow => {
                conditional_format.criteria = libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_TOMORROW as u8;
            }
            ConditionalFormatTimePeriodCriteria::Last7Days => {
                conditional_format.criteria = libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_7_DAYS as u8;
            }
            ConditionalFormatTimePeriodCriteria::LastWeek => {
                conditional_format.criteria = libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_WEEK as u8;
            }
            ConditionalFormatTimePeriodCriteria::ThisWeek => {
                conditional_format.criteria = libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_THIS_WEEK as u8;
            }
            ConditionalFormatTimePeriodCriteria::NextWeek => {
                conditional_format.criteria = libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_NEXT_WEEK as u8;
            }
            ConditionalFormatTimePeriodCriteria::LastMonth => {
                conditional_format.criteria = libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_MONTH as u8;
            }
            ConditionalFormatTimePeriodCriteria::ThisMonth => {
                conditional_format.criteria = libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_THIS_MONTH as u8;
            }
            ConditionalFormatTimePeriodCriteria::NextMonth => {
                conditional_format.criteria = libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_NEXT_MONTH as u8;
            }
        }
        Ok(())
    }
}

impl ConditionalFormat {
    /// This function used to specify Excel's "Dates Occurring" style conditional format.
    ///
    /// See [`ConditionalFormatTimePeriodCriteria`] to learn available criteria.
    ///
    /// ```rust
    /// # use xlsxwriter::prelude::*;
    /// # use xlsxwriter::worksheet::conditional_format::*;
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-conditional_format-time_period.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # worksheet.write_datetime(0, 0, &DateTime::date(2022, 12, 4), Some(Format::new().set_num_format("yyyy/m/d")))?;
    /// worksheet.conditional_format_cell(
    ///     0, 0,
    ///     &ConditionalFormat::time_period(
    ///         ConditionalFormatTimePeriodCriteria::Today,
    ///         Format::new().set_bg_color(FormatColor::Yellow),
    ///     ),
    /// )?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn time_period(
        time_period: ConditionalFormatTimePeriodCriteria,
        format: &Format,
    ) -> ConditionalFormat {
        ConditionalFormat::ConditionType {
            criteria: ConditionalFormatTypes::TimePeriod(time_period),
            format: format.clone(),
        }
    }
}

#[cfg(test)]
mod test {
    use chrono::Months;

    use crate::*;

    #[cfg(feature = "chrono")]
    #[test]
    fn test_worksheet_conditional_format2() -> Result<(), XlsxError> {
        use crate::{conditional_format::*, DateTime};
        use chrono::Days;
        use std::convert::TryInto;

        let workbook = Workbook::new("test-worksheet_conditional-format2.xlsx")?;
        let mut worksheet = workbook.add_worksheet(Some("Conditional format"))?;

        let today = chrono::Local::now().naive_local().date();
        let two_month_ago = today.checked_sub_months(Months::new(2)).unwrap();
        let last_month = today.checked_sub_months(Months::new(1)).unwrap();

        let mut date_format = Format::new();
        date_format.set_num_format("yyyy/m/d (ddd)");
        let mut yellow_background = Format::new();
        yellow_background.set_bg_color(FormatColor::Yellow);

        worksheet.write_string(1, 0, "Two month ago", None)?;
        for i in 1..11 {
            worksheet.write_datetime(1, i, &two_month_ago.into(), Some(&date_format))?;
        }

        worksheet.write_string(2, 0, "Last month", None)?;
        for i in 1..11 {
            worksheet.write_datetime(2, i, &last_month.into(), Some(&date_format))?;
        }

        for d in 0..=30 {
            let day: DateTime = if d <= 15 {
                worksheet.write_string(d + 3, 0, &format!("Today - {}", 15 - d), None)?;
                today
                    .checked_sub_days(Days::new((15 - d).try_into().unwrap()))
                    .unwrap()
            } else {
                worksheet.write_string(d + 3, 0, &format!("Today + {}", d - 15), None)?;
                today.checked_add_days(Days::new((d - 15).into())).unwrap()
            }
            .into();
            for i in 1..11 {
                worksheet.write_datetime(d + 3, i, &day, Some(&date_format))?;
            }
        }

        let next_month = today.checked_add_months(Months::new(1)).unwrap();
        let two_month_after = today.checked_add_months(Months::new(2)).unwrap();

        worksheet.write_string(34, 0, "Next month", None)?;
        for i in 1..11 {
            worksheet.write_datetime(34, i, &next_month.into(), Some(&date_format))?;
        }

        worksheet.write_string(35, 0, "Two month after", None)?;
        for i in 1..11 {
            worksheet.write_datetime(35, i, &two_month_after.into(), Some(&date_format))?;
        }

        worksheet.set_column(0, 10, 20.0, None)?;

        worksheet.write_string(0, 1, "Today", None)?;
        worksheet.conditional_format_range(
            1,
            1,
            35,
            1,
            &ConditionalFormat::time_period(
                ConditionalFormatTimePeriodCriteria::Today,
                &yellow_background,
            ),
        )?;

        worksheet.write_string(0, 2, "Yesterday", None)?;
        worksheet.conditional_format_range(
            1,
            2,
            35,
            2,
            &ConditionalFormat::time_period(
                ConditionalFormatTimePeriodCriteria::Yesterday,
                &yellow_background,
            ),
        )?;

        worksheet.write_string(0, 3, "Tomorrow", None)?;
        worksheet.conditional_format_range(
            1,
            3,
            35,
            3,
            &ConditionalFormat::time_period(
                ConditionalFormatTimePeriodCriteria::Tomorrow,
                &yellow_background,
            ),
        )?;

        worksheet.write_string(0, 4, "Last 7 days", None)?;
        worksheet.conditional_format_range(
            1,
            4,
            35,
            4,
            &ConditionalFormat::time_period(
                ConditionalFormatTimePeriodCriteria::Last7Days,
                &yellow_background,
            ),
        )?;

        worksheet.write_string(0, 5, "Last week", None)?;
        worksheet.conditional_format_range(
            1,
            5,
            35,
            5,
            &ConditionalFormat::time_period(
                ConditionalFormatTimePeriodCriteria::LastWeek,
                &yellow_background,
            ),
        )?;

        worksheet.write_string(0, 6, "This week", None)?;
        worksheet.conditional_format_range(
            1,
            6,
            35,
            6,
            &ConditionalFormat::time_period(
                ConditionalFormatTimePeriodCriteria::ThisWeek,
                &yellow_background,
            ),
        )?;

        worksheet.write_string(0, 7, "Next week", None)?;
        worksheet.conditional_format_range(
            1,
            7,
            35,
            7,
            &ConditionalFormat::time_period(
                ConditionalFormatTimePeriodCriteria::NextWeek,
                &yellow_background,
            ),
        )?;

        worksheet.write_string(0, 8, "Last month", None)?;
        worksheet.conditional_format_range(
            1,
            8,
            35,
            8,
            &ConditionalFormat::time_period(
                ConditionalFormatTimePeriodCriteria::LastMonth,
                &yellow_background,
            ),
        )?;

        worksheet.write_string(0, 9, "This month", None)?;
        worksheet.conditional_format_range(
            1,
            9,
            35,
            9,
            &ConditionalFormat::time_period(
                ConditionalFormatTimePeriodCriteria::ThisMonth,
                &yellow_background,
            ),
        )?;

        worksheet.write_string(0, 10, "Next month", None)?;
        worksheet.conditional_format_range(
            1,
            10,
            35,
            10,
            &ConditionalFormat::time_period(
                ConditionalFormatTimePeriodCriteria::NextMonth,
                &yellow_background,
            ),
        )?;

        Ok(())
    }
}
