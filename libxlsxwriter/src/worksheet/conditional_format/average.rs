use crate::XlsxError;

/// The Average type is used to specify Excel's "Average" style conditional format.
///
/// See [`super::ConditionalFormat::average`] to learn usage.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub enum ConditionalFormatAverageCriteria {
    /// Format cells above the average for the range.     
    AverageAbove,
    /// Format cells below the average for the range.
    AverageBelow,
    /// Format cells above or equal to the average for the range.
    AverageAboveOrEqual,
    /// Format cells below or equal to the average for the range.
    AverageBelowOrEqual,
    /// Format cells 1 standard deviation above the average for the range.
    Average1StdDevAbove,
    /// Format cells 1 standard deviation below the average for the range.
    Average1StdDevBelow,
    /// Format cells 2 standard deviation above the average for the range.
    Average2StdDevAbove,
    /// Format cells 2 standard deviation below the average for the range.
    Average2StdDevBelow,
    /// Format cells 3 standard deviation above the average for the range.
    Average3StdDevAbove,
    /// Format cells 3 standard deviation below the average for the range.
    Average3StdDevBelow,
}

impl ConditionalFormatAverageCriteria {
    pub(crate) fn into_internal_value(
        &self,
        conditional_format: &mut libxlsxwriter_sys::lxw_conditional_format,
    ) -> Result<(), XlsxError> {
        conditional_format.type_ =
            libxlsxwriter_sys::lxw_conditional_format_types_LXW_CONDITIONAL_TYPE_AVERAGE as u8;
        match self {
            ConditionalFormatAverageCriteria::AverageAbove => {
                conditional_format.criteria =  libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_AVERAGE_ABOVE as u8;
            }
            ConditionalFormatAverageCriteria::AverageBelow => {
                conditional_format.criteria = libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_AVERAGE_BELOW as u8;
            }
            ConditionalFormatAverageCriteria::AverageAboveOrEqual => {
                conditional_format.criteria = libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_AVERAGE_ABOVE_OR_EQUAL as u8;
            }
            ConditionalFormatAverageCriteria::AverageBelowOrEqual => {
                conditional_format.criteria = libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_AVERAGE_BELOW_OR_EQUAL as u8;
            }
            ConditionalFormatAverageCriteria::Average1StdDevAbove => {
                conditional_format.criteria = libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_AVERAGE_1_STD_DEV_ABOVE as u8;
            }
            ConditionalFormatAverageCriteria::Average1StdDevBelow => {
                conditional_format.criteria = libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_AVERAGE_1_STD_DEV_BELOW as u8;
            }
            ConditionalFormatAverageCriteria::Average2StdDevAbove => {
                conditional_format.criteria = libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_AVERAGE_2_STD_DEV_ABOVE as u8;
            }
            ConditionalFormatAverageCriteria::Average2StdDevBelow => {
                conditional_format.criteria = libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_AVERAGE_2_STD_DEV_BELOW as u8;
            }
            ConditionalFormatAverageCriteria::Average3StdDevAbove => {
                conditional_format.criteria = libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_AVERAGE_3_STD_DEV_ABOVE as u8;
            }
            ConditionalFormatAverageCriteria::Average3StdDevBelow => {
                conditional_format.criteria = libxlsxwriter_sys::lxw_conditional_criteria_LXW_CONDITIONAL_CRITERIA_AVERAGE_3_STD_DEV_BELOW as u8;
            }
        }
        Ok(())
    }
}

#[cfg(test)]
mod test {
    use crate::worksheet::conditional_format::*;
    use crate::*;

    #[test]
    fn test_worksheet_conditional_format3() -> Result<(), XlsxError> {
        let workbook = Workbook::new("test-worksheet_conditional-format3.xlsx")?;
        let mut worksheet = workbook.add_worksheet(Some("Conditional format"))?;
        for i in 0..=100 {
            for j in 0..10 {
                worksheet.write_number(i + 1, j, i.into(), None)?;
            }
        }

        worksheet.write_string(0, 0, "Average Above", None)?;
        worksheet.conditional_format_range(
            1,
            0,
            102,
            0,
            &ConditionalFormat::average(
                ConditionalFormatAverageCriteria::AverageAbove,
                Format::new().set_bg_color(FormatColor::Yellow),
            ),
        )?;

        worksheet.write_string(0, 1, "Average Below", None)?;
        worksheet.conditional_format_range(
            1,
            1,
            102,
            1,
            &ConditionalFormat::average(
                ConditionalFormatAverageCriteria::AverageBelow,
                Format::new().set_bg_color(FormatColor::Yellow),
            ),
        )?;

        worksheet.write_string(0, 2, "Average Above or Equal", None)?;
        worksheet.conditional_format_range(
            1,
            2,
            102,
            2,
            &ConditionalFormat::average(
                ConditionalFormatAverageCriteria::AverageAboveOrEqual,
                Format::new().set_bg_color(FormatColor::Yellow),
            ),
        )?;

        worksheet.write_string(0, 3, "Average Below or Equal", None)?;
        worksheet.conditional_format_range(
            1,
            3,
            102,
            3,
            &ConditionalFormat::average(
                ConditionalFormatAverageCriteria::AverageBelowOrEqual,
                Format::new().set_bg_color(FormatColor::Yellow),
            ),
        )?;

        worksheet.write_string(0, 4, "Average 1 Standard deviation above", None)?;
        worksheet.conditional_format_range(
            1,
            4,
            102,
            4,
            &ConditionalFormat::average(
                ConditionalFormatAverageCriteria::Average1StdDevAbove,
                Format::new().set_bg_color(FormatColor::Yellow),
            ),
        )?;

        worksheet.write_string(0, 5, "Average 1 Standard deviation below", None)?;
        worksheet.conditional_format_range(
            1,
            5,
            102,
            5,
            &ConditionalFormat::average(
                ConditionalFormatAverageCriteria::Average1StdDevBelow,
                Format::new().set_bg_color(FormatColor::Yellow),
            ),
        )?;

        worksheet.write_string(0, 6, "Average 2 Standard deviation above", None)?;
        worksheet.conditional_format_range(
            1,
            6,
            102,
            6,
            &ConditionalFormat::average(
                ConditionalFormatAverageCriteria::Average2StdDevAbove,
                Format::new().set_bg_color(FormatColor::Yellow),
            ),
        )?;

        worksheet.write_string(0, 7, "Average 2 Standard deviation below", None)?;
        worksheet.conditional_format_range(
            1,
            7,
            102,
            7,
            &ConditionalFormat::average(
                ConditionalFormatAverageCriteria::Average2StdDevBelow,
                Format::new().set_bg_color(FormatColor::Yellow),
            ),
        )?;

        worksheet.write_string(0, 8, "Average 3 Standard deviation above", None)?;
        worksheet.conditional_format_range(
            1,
            8,
            102,
            8,
            &ConditionalFormat::average(
                ConditionalFormatAverageCriteria::Average3StdDevAbove,
                Format::new().set_bg_color(FormatColor::Yellow),
            ),
        )?;

        worksheet.write_string(0, 9, "Average 3 Standard deviation below", None)?;
        worksheet.conditional_format_range(
            1,
            9,
            102,
            9,
            &ConditionalFormat::average(
                ConditionalFormatAverageCriteria::Average3StdDevBelow,
                Format::new().set_bg_color(FormatColor::Yellow),
            ),
        )?;
        Ok(())
    }
}
