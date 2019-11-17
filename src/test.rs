use super::*;

#[test]
fn test_sample1() -> Result<(), XlsxError> {
    let workbook = Workbook::new("target/test.xlsx");

    let mut format1 = workbook.get_format();
    format1.set_bold();
    format1.set_font_name("Arial");
    format1.set_font_color(FormatColor::Red);
    format1.set_italic();
    format1.set_underline(FormatUnderline::Single);

    let mut format2 = workbook.get_format();
    format2.set_font_color(FormatColor::Blue);
    format2.set_underline(FormatUnderline::Double);

    let mut format3 = workbook.get_format();
    format3.set_font_color(FormatColor::Blue);
    format3.set_underline(FormatUnderline::Single);

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
