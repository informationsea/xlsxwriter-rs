use super::*;

#[test]
fn test_simple1() -> Result<(), XlsxError> {
    let workbook = Workbook::new("target/simple1.xlsx");
    let mut format1 = workbook.get_format();
    format1.set_font_color(FormatColor::Red);

    let mut format2 = workbook.get_format();
    format2
        .set_font_color(FormatColor::Blue)
        .set_underline(FormatUnderline::Single);

    let mut format3 = workbook.get_format();
    format3
        .set_font_color(FormatColor::Green)
        .set_align(FormatAlignment::CenterAcross)
        .set_align(FormatAlignment::VerticalCenter);

    let mut sheet1 = workbook.add_worksheet(None)?;
    sheet1.write_string(0, 0, "Red text", Some(&format1))?;
    sheet1.write_number(0, 1, 20., None)?;
    sheet1.write_formula_num(1, 0, "=10+B1", None, 30.)?;
    sheet1.write_url(
        1,
        1,
        "https://github.com/informationsea/xlsxwriter-rs",
        Some(&format2),
    )?;
    sheet1.merge_range(2, 0, 3, 2, "Hello, world", Some(&format3))?;

    sheet1.set_selection(1, 0, 1, 2);
    sheet1.set_tab_color(FormatColor::Cyan);
    workbook.close()?;

    Ok(())
}

#[test]
fn test_sample1() -> Result<(), XlsxError> {
    let workbook = Workbook::new("target/test.xlsx");

    let mut format1 = workbook.get_format();
    format1
        .set_bold()
        .set_font_name("Arial")
        .set_font_color(FormatColor::Red)
        .set_italic()
        .set_underline(FormatUnderline::Single);

    let mut format2 = workbook.get_format();
    format2
        .set_font_color(FormatColor::Blue)
        .set_underline(FormatUnderline::Double);

    let mut format3 = workbook.get_format();
    format3
        .set_font_color(FormatColor::Blue)
        .set_underline(FormatUnderline::Single);

    let mut format4 = workbook.get_format();
    format4.set_num_format("mmm d yyyy hh:mm AM/PM");

    let mut sheet = workbook.add_worksheet(None)?;
    sheet.write_string(0, 0, "Hello", Some(&format1))?;
    sheet.write_formula(1, 0, "=1+2", None)?;
    sheet.write_rich_string(2, 0, &[("hello", &format1), (" world", &format2)], None)?;
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
