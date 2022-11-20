use super::*;

#[test]
fn test_simple1() -> Result<(), XlsxError> {
    let workbook = Workbook::new("../target/simple1.xlsx")?;

    let mut sheet1 = workbook.add_worksheet(None)?;
    sheet1.write_string(
        0,
        0,
        "Red text",
        Some(&Format::new().set_font_color(FormatColor::Red)),
    )?;
    sheet1.write_number(0, 1, 20., None)?;
    sheet1.write_formula_num(1, 0, "=10+B1", None, 30.)?;
    sheet1.write_url(
        1,
        1,
        "https://github.com/informationsea/xlsxwriter-rs",
        Some(
            &Format::new()
                .set_font_color(FormatColor::Blue)
                .set_underline(FormatUnderline::Single),
        ),
    )?;
    sheet1.merge_range(
        2,
        0,
        3,
        2,
        "Hello, world",
        Some(
            &Format::new()
                .set_font_color(FormatColor::Green)
                .set_align(FormatAlignment::CenterAcross)
                .set_vertical_align(FormatVerticalAlignment::VerticalCenter),
        ),
    )?;

    sheet1.set_selection(1, 0, 1, 2);
    sheet1.set_tab_color(FormatColor::Cyan);
    workbook.close()?;

    Ok(())
}

#[test]
fn test_sample1() -> Result<(), XlsxError> {
    let workbook = Workbook::new("../target/test.xlsx")?;

    let mut format1 = workbook.add_format();
    format1
        .set_bold()
        .set_font_name("Arial")
        .set_font_color(FormatColor::Red)
        .set_italic()
        .set_underline(FormatUnderline::Single);

    let mut format2 = workbook.add_format();
    format2
        .set_font_color(FormatColor::Blue)
        .set_underline(FormatUnderline::Double);

    let mut format3 = workbook.add_format();
    format3
        .set_font_color(FormatColor::Blue)
        .set_underline(FormatUnderline::Single);

    let mut format4 = workbook.add_format();
    format4.set_num_format("mmm d yyyy hh:mm AM/PM");

    let mut sheet = workbook.add_worksheet(None)?;
    sheet.write_string(0, 0, "Hello", Some(&format1))?;
    sheet.write_formula(1, 0, "=1+2", None)?;
    sheet.write_rich_string(
        2,
        0,
        &[("hello", Some(&format1)), (" world", Some(&format2))],
        None,
    )?;
    sheet.write_number(3, 0, 100., None)?;
    sheet.write_formula_num(4, 0, "=3+2", None, 5.)?;
    sheet.write_url(0, 1, "https://github.com", Some(&format3))?;
    sheet.write_blank(1, 1, None)?;
    sheet.write_datetime(
        2,
        1,
        &DateTime::new(2019, 11, 12, 3, 8, 10.),
        Some(&format4),
    )?;
    workbook.close()?;
    Ok(())
}

#[test]
fn test_add_table1() -> Result<(), XlsxError> {
    let workbook = Workbook::new("test-worksheet_add_table-2.xlsx")?;
    let mut worksheet = workbook.add_worksheet(None)?;
    worksheet.write_string(0, 0, "header 1", None)?;
    worksheet.write_string(0, 1, "header 2", None)?;
    worksheet.write_string(1, 0, "content 1", None)?;
    worksheet.write_number(1, 1, 1.0, None)?;
    worksheet.write_string(2, 0, "content 2", None)?;
    worksheet.write_number(2, 1, 2.0, None)?;
    worksheet.write_string(3, 0, "content 3", None)?;
    worksheet.write_number(3, 1, 3.0, None)?;

    let options = TableOptions {
        name: Some("TABLE1".to_string()),
        no_header_row: true,
        no_autofilter: true,
        no_banded_rows: true,
        banded_columns: true,
        first_column: true,
        last_column: true,
        style_type: TableStyleType::Light,
        style_type_number: 1,
        total_row: false,
        columns: None,
    };
    worksheet.add_table(0, 0, 3, 1, Some(options))?;
    workbook.close()?;
    Ok(())
}

#[test]
fn test_add_table2() -> Result<(), XlsxError> {
    let workbook = Workbook::new("test-worksheet_add_table-3.xlsx")?;
    let mut worksheet = workbook.add_worksheet(None)?;
    worksheet.write_string(0, 0, "header 1", None)?;
    worksheet.write_string(0, 1, "header 2", None)?;
    worksheet.write_string(0, 2, "header 3", None)?;
    worksheet.write_string(1, 0, "content 1", None)?;
    worksheet.write_number(1, 1, 1.0, None)?;
    worksheet.write_string(2, 0, "content 2", None)?;
    worksheet.write_number(2, 1, 2.0, None)?;
    worksheet.write_string(3, 0, "content 3", None)?;
    worksheet.write_number(3, 1, 3.0, None)?;

    let options = TableOptions {
        name: Some("TABLE1".to_string()),
        no_header_row: false,
        no_autofilter: false,
        no_banded_rows: false,
        banded_columns: false,
        first_column: false,
        last_column: false,
        style_type: TableStyleType::Dark,
        style_type_number: 5,
        total_row: true,
        columns: Some(vec![
            TableColumn {
                header: Some("HEADER1".to_string()),
                formula: None,
                total_string: Some("Total".to_string()),
                total_function: TableTotalFunction::None,
                header_format: None,
                format: None,
                total_value: 0.0,
            },
            TableColumn {
                header: Some("HEADER2".to_string()),
                formula: None,
                total_string: None,
                total_function: TableTotalFunction::Sum,
                header_format: None,
                format: None,
                total_value: 6.0,
            },
            TableColumn {
                header: Some("HEADER3".to_string()),
                formula: Some("=[@[HEADER2]]*2".to_string()),
                total_string: None,
                total_function: TableTotalFunction::Count,
                header_format: None,
                format: None,
                total_value: 3.0,
            },
        ]),
    };
    worksheet.add_table(0, 0, 4, 2, Some(options))?;
    workbook.close()?;
    Ok(())
}
