use crate::excel::data::days::Day;
use anyhow::Result as AResult;
use rust_xlsxwriter::{Format, FormatAlign, utility::column_name_to_number, worksheet::Worksheet};

pub(crate) fn add_header_cells(month_worksheet: &mut Worksheet, first_day: &Day) -> AResult<()> {
    // Year
    month_worksheet.write(0, column_name_to_number("A"), first_day.year())?;
    // AppInfo
    let appinfo_format = Format::new().set_align(FormatAlign::Center);
    month_worksheet.merge_range(
        0,
        column_name_to_number("B"),
        0,
        column_name_to_number("C"),
        "dev.release",
        &appinfo_format,
    )?;

    // Month
    let month_format = Format::new().set_align(FormatAlign::Center);
    month_worksheet.merge_range(
        1,
        column_name_to_number("A"),
        1,
        column_name_to_number("C"),
        first_day.month_name().as_str(),
        &month_format,
    )?;
    // Day header
    month_worksheet.write(2, column_name_to_number("A"), "Число/День")?;
    // Hours header
    month_worksheet.write(2, column_name_to_number("B"), "Часы")?;
    // Bonus header
    month_worksheet.write(2, column_name_to_number("C"), "Доплата")?;

    // Total month work hours header
    month_worksheet.write(0, column_name_to_number("D"), "Рабочие часы:")?;
    // Total weekends hours header
    month_worksheet.write(1, column_name_to_number("D"), "Часы выходных:")?;
    // Total overvork hours header
    month_worksheet.write(2, column_name_to_number("D"), "Часы переработки:")?;

    // Salary input header
    month_worksheet.write(4, column_name_to_number("D"), "Оклад:")?;
    // Total payout header
    month_worksheet.write(5, column_name_to_number("D"), "К получению:")?;
    Ok(())
}
