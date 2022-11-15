use crate::{
    try_to_vec, CStringHelper, StringOrFloat, Worksheet, WorksheetCol, WorksheetRow, XlsxError,
};

/// And/or operator conditions when using 2 filter rules with `filter_column2`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum FilterOperator {
    FilterAnd,
    FilterOr,
}

impl Default for FilterOperator {
    fn default() -> Self {
        FilterOperator::FilterAnd
    }
}

impl FilterOperator {
    pub(crate) fn into_internal(self) -> libxlsxwriter_sys::lxw_filter_operator {
        match self {
            FilterOperator::FilterAnd => libxlsxwriter_sys::lxw_filter_operator_LXW_FILTER_AND,
            FilterOperator::FilterOr => libxlsxwriter_sys::lxw_filter_operator_LXW_FILTER_OR,
        }
    }
}

/// Criteria used to define an autofilter rule condition.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum FilterCriteria {
    /// Filter cells equal to a value.
    EqualTo,
    /// Filter cells not equal to a value.
    NotEqualTo,
    /// Filter cells greater than a value.
    GreaterThan,
    /// Filter cells less than a value.
    LessThan,
    /// Filter cells greater than or equal to a value.
    GreaterThanOrEqualTo,
    /// Filter cells less than or equal to a value.
    LessThanOrEqualTo,
    /// Filter cells that are blank.
    Blanks,
    /// Filter cells that are not blank.
    NonBlanks,
}

impl Default for FilterCriteria {
    fn default() -> Self {
        FilterCriteria::EqualTo
    }
}

impl FilterCriteria {
    pub(crate) fn into_internal(self) -> libxlsxwriter_sys::lxw_filter_criteria {
        match self {
            FilterCriteria::EqualTo => {
                libxlsxwriter_sys::lxw_filter_criteria_LXW_FILTER_CRITERIA_EQUAL_TO
            }
            FilterCriteria::NotEqualTo => {
                libxlsxwriter_sys::lxw_filter_criteria_LXW_FILTER_CRITERIA_NOT_EQUAL_TO
            }
            FilterCriteria::GreaterThan => {
                libxlsxwriter_sys::lxw_filter_criteria_LXW_FILTER_CRITERIA_GREATER_THAN
            }
            FilterCriteria::LessThan => {
                libxlsxwriter_sys::lxw_filter_criteria_LXW_FILTER_CRITERIA_LESS_THAN
            }
            FilterCriteria::GreaterThanOrEqualTo => {
                libxlsxwriter_sys::lxw_filter_criteria_LXW_FILTER_CRITERIA_GREATER_THAN_OR_EQUAL_TO
            }
            FilterCriteria::LessThanOrEqualTo => {
                libxlsxwriter_sys::lxw_filter_criteria_LXW_FILTER_CRITERIA_LESS_THAN_OR_EQUAL_TO
            }
            FilterCriteria::Blanks => {
                libxlsxwriter_sys::lxw_filter_criteria_LXW_FILTER_CRITERIA_BLANKS
            }
            FilterCriteria::NonBlanks => {
                libxlsxwriter_sys::lxw_filter_criteria_LXW_FILTER_CRITERIA_NON_BLANKS
            }
        }
    }
}

/// Options for autofilter rules.
#[derive(Debug, Clone, PartialEq, PartialOrd, Default)]
pub struct FilterRule {
    /// The FilterCriteria to define the rule.
    pub criteria: FilterCriteria,
    /// value to which the criteria applies.
    pub value: StringOrFloat,
}

impl FilterRule {
    pub fn new<T: Into<StringOrFloat>>(criteria: FilterCriteria, value: T) -> Self {
        FilterRule {
            criteria,
            value: value.into(),
        }
    }

    pub(crate) fn into_internal(
        &self,
        c_string_helper: &mut CStringHelper,
    ) -> Result<libxlsxwriter_sys::lxw_filter_rule, XlsxError> {
        Ok(libxlsxwriter_sys::lxw_filter_rule {
            criteria: self.criteria.into_internal() as u8,
            value_string: c_string_helper.add_opt(self.value.to_str())? as *mut std::os::raw::c_char,
            value: self.value.to_f64().unwrap_or_default(),
        })
    }
}

impl<'a> Worksheet<'a> {
    /// This function allows an autofilter to be added to a worksheet.
    ///
    /// An autofilter is a way of adding drop down lists to the headers of a 2D range of worksheet data.
    /// This allows users to filter the data based on simple criteria so that some data is shown and some is hidden.
    ///
    pub fn autofilter(
        &mut self,
        first_row: WorksheetRow,
        first_col: WorksheetCol,
        last_row: WorksheetRow,
        last_col: WorksheetCol,
    ) -> Result<(), XlsxError> {
        unsafe {
            let result = libxlsxwriter_sys::worksheet_autofilter(
                self.worksheet,
                first_row,
                first_col,
                last_row,
                last_col,
            );
            if result == libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                Ok(())
            } else {
                Err(XlsxError::new(result))
            }
        }
    }

    /// This function can be used to filter columns in a autofilter range based on single rule conditions.
    ///
    /// ```rust
    /// use xlsxwriter::*;
    ///
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-worksheet_filter_column.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # worksheet.write_string(0, 0, "Column A", None)?;
    /// # worksheet.write_string(0, 1, "Column B", None)?;
    /// # for i in 1..=10 {
    /// #     let v: f64 = i.into();
    /// #     worksheet.write_number(i, 0, v, None)?;
    /// #     worksheet.write_number(i, 1, v + 0.5, None)?;
    /// # }
    /// worksheet.autofilter(0, 0, 10, 1)?;
    /// worksheet.filter_column(0, &FilterRule::new(FilterCriteria::NotEqualTo, 9.0))?;
    /// worksheet.set_row_opt(9, LXW_DEF_ROW_HEIGHT, None, &RowColOptions::new(true, 0, false))?;
    /// # Ok(())
    /// # }
    ///
    /// ```
    ///
    /// The `col` parameter is a zero indexed column number and must refer to a column in an existing autofilter created with [`Worksheet::autofilter`].
    /// It isn't sufficient to just specify the filter condition. You must also hide any rows that don't match the filter condition.
    pub fn filter_column(
        &mut self,
        col: crate::WorksheetCol,
        rule: &FilterRule,
    ) -> Result<(), XlsxError> {
        unsafe {
            let mut c_string_helper = CStringHelper::new();
            let mut rule_converted = rule.into_internal(&mut c_string_helper)?;
            let e = libxlsxwriter_sys::worksheet_filter_column(
                self.worksheet,
                col,
                &mut rule_converted,
            );
            if e != libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                return Err(XlsxError::new(e));
            }
        }
        Ok(())
    }

    /// This function can be used to filter columns in a autofilter range based on two rule conditions.
    ///
    /// ```rust
    /// use xlsxwriter::*;
    ///
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-worksheet_filter_column2.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # worksheet.write_string(0, 0, "Column A", None)?;
    /// # worksheet.write_string(0, 1, "Column B", None)?;
    /// # for i in 1..=10 {
    /// #     let v: f64 = i.into();
    /// #     worksheet.write_number(i, 0, v, None)?;
    /// #     worksheet.write_number(i, 1, v + 0.5, None)?;
    /// # }
    /// worksheet.autofilter(0, 0, 10, 1)?;
    /// worksheet.filter_column2(0, &FilterRule::new(FilterCriteria::LessThan, 8.5), &FilterRule::new(FilterCriteria::GreaterThan, 9.5), FilterOperator::FilterOr)?;
    /// worksheet.set_row_opt(9, LXW_DEF_ROW_HEIGHT, None, &RowColOptions::new(true, 0, false))?;
    /// # Ok(())
    /// # }
    ///
    /// ```
    ///
    /// It isn't sufficient to just specify the filter condition. You must also hide any rows that don't match the filter condition.
    pub fn filter_column2(
        &mut self,
        col: crate::WorksheetCol,
        rule1: &FilterRule,
        rule2: &FilterRule,
        and_or: FilterOperator,
    ) -> Result<(), XlsxError> {
        unsafe {
            let mut c_string_helper = CStringHelper::new();
            let mut rule_converted1 = rule1.into_internal(&mut c_string_helper)?;
            let mut rule_converted2 = rule2.into_internal(&mut c_string_helper)?;
            let e = libxlsxwriter_sys::worksheet_filter_column2(
                self.worksheet,
                col,
                &mut rule_converted1,
                &mut rule_converted2,
                and_or.into_internal() as u8,
            );
            if e != libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                return Err(XlsxError::new(e));
            }
        }
        Ok(())
    }

    /// This function can be used specify multiple string matching criteria. This is a newer type of filter introduced in Excel 2007.
    /// Prior to that it was only possible to have either 1 or 2 filter conditions, such as the ones used by [`Worksheet::filter_column()`]
    /// and [`Worksheet::filter_column2()`].
    ///
    /// ```rust
    /// use xlsxwriter::*;
    ///
    /// # fn main() -> Result<(), XlsxError> {
    /// # let workbook = Workbook::new("test-worksheet_filter_list.xlsx")?;
    /// # let mut worksheet = workbook.add_worksheet(None)?;
    /// # worksheet.write_string(0, 0, "Column A", None)?;
    /// # worksheet.write_string(0, 1, "Column B", None)?;
    /// # for i in 1..=10 {
    /// #     let v: f64 = i.into();
    /// #     worksheet.write_number(i, 0, v, None)?;
    /// #     worksheet.write_number(i, 1, v + 0.5, None)?;
    /// # }
    /// worksheet.autofilter(0, 0, 10, 1)?;
    /// worksheet.filter_list(0, &["1", "2", "3", "4", "5", "6", "7", "8", "10"])?;
    /// worksheet.set_row_opt(9, LXW_DEF_ROW_HEIGHT, None, &RowColOptions::new(true, 0, false))?;
    /// # Ok(())
    /// # }
    ///
    /// ```
    ///
    /// It isn't sufficient to just specify the filter condition. You must also hide any rows that don't match the filter condition.    
    pub fn filter_list(
        &mut self,
        col: crate::WorksheetCol,
        list: &[&str],
    ) -> Result<(), XlsxError> {
        let mut cstring_helper = crate::CStringHelper::new();
        let mut cstr_list: Vec<_> = try_to_vec(
            list.iter()
                .map(|x| Ok(cstring_helper.add(x)? as *mut std::os::raw::c_char)),
        )?;
        cstr_list.push(std::ptr::null_mut());
        unsafe {
            let result = libxlsxwriter_sys::worksheet_filter_list(
                self.worksheet,
                col,
                cstr_list.as_mut_ptr(),
            );
            if result != libxlsxwriter_sys::lxw_error_LXW_NO_ERROR {
                return Err(XlsxError::new(result));
            }
        }
        Ok(())
    }
}

#[cfg(test)]
mod test {
    use crate::{RowColOptions, Workbook};

    use super::*;

    fn create_sheet(worksheet: &mut Worksheet) -> Result<(), XlsxError> {
        worksheet.write_string(0, 0, "Column A", None)?;
        worksheet.write_string(0, 1, "Column B", None)?;
        worksheet.write_string(0, 2, "Column C", None)?;
        worksheet.write_string(0, 3, "Column D", None)?;

        for i in 1..20 {
            let value: f64 = i.into();
            worksheet.write_number(i, 0, value, None)?;
            worksheet.write_number(i, 1, value / 2., None)?;
            worksheet.write_string(i, 2, &format!("TEXT{:02}", i), None)?;
            worksheet.write_boolean(i, 3, i % 2 == 0, None)?;
        }
        worksheet.autofilter(0, 0, 19, 3)?;
        Ok(())
    }

    #[test]
    fn test_autofilter() -> Result<(), XlsxError> {
        let workbook = Workbook::new("test-worksheet_autofilter.xlsx")?;
        // ------------
        let mut worksheet1 = workbook.add_worksheet(Some("Sheet 1"))?;
        create_sheet(&mut worksheet1)?;
        let worksheet1_criteria = FilterRule::new(FilterCriteria::GreaterThan, 10.0);
        worksheet1.filter_column(0, &worksheet1_criteria)?;
        let mut hidden_row = RowColOptions::new(true, 0, false);
        for i in 1..=10 {
            worksheet1.set_row_opt(i, 13.2, None, &mut hidden_row)?;
        }
        // ------------
        let mut worksheet2 = workbook.add_worksheet(Some("Sheet 2"))?;
        create_sheet(&mut worksheet2)?;
        let worksheet2_criteria1 = FilterRule::new(FilterCriteria::GreaterThanOrEqualTo, 3.0);
        let worksheet2_criteria2 = FilterRule::new(FilterCriteria::LessThan, 5.5);
        worksheet2.filter_column2(
            1,
            &worksheet2_criteria1,
            &worksheet2_criteria2,
            FilterOperator::FilterAnd,
        )?;
        for i in 1..20 {
            let value: f64 = i.into();
            if (value / 2.0) < 3. || 5.5 <= (value / 2.0) {
                worksheet2.set_row_opt(i, 13.2, None, &mut hidden_row)?;
            }
        }
        // --------------
        let mut worksheet3 = workbook.add_worksheet(Some("Sheet 3"))?;
        create_sheet(&mut worksheet3)?;
        worksheet3.filter_list(2, &["TEXT02", "TEXT03"])?;
        for i in 1..20 {
            if i != 2 && i != 3 {
                worksheet3.set_row_opt(i, 13.2, None, &mut hidden_row)?;
            }
        }

        workbook.close()?;
        Ok(())
    }
}
