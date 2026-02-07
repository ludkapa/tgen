use crate::excel::{
    data::days::{Day, Season},
    design::{CellType, DataType, cell_style},
};
use anyhow::Result as AResult;
use rust_xlsxwriter::{utility::column_name_to_number, worksheet::Worksheet};

pub(crate) fn add_header_cells(month_worksheet: &mut Worksheet, first_day: &Day) -> AResult<()> {
    // Year
    let mut format = cell_style(DataType::UsualText, CellType::Header);
    month_worksheet.write_with_format(0, column_name_to_number("A"), first_day.year(), &format)?;
    // AppInfo
    month_worksheet.merge_range(
        0,
        column_name_to_number("B"),
        0,
        column_name_to_number("C"),
        "dev.release",
        &format,
    )?;

    // Day header
    month_worksheet.write_with_format(2, column_name_to_number("A"), "Число/День", &format)?;
    // Bonus header
    month_worksheet.write_with_format(2, column_name_to_number("C"), "Доплата", &format)?;
    // Total month work hours header
    month_worksheet.write_with_format(0, column_name_to_number("D"), "Рабочие часы:", &format)?;
    // Total weekends hours header
    month_worksheet.write_with_format(1, column_name_to_number("D"), "Часы выходных:", &format)?;
    // Total overvork hours header
    month_worksheet.write_with_format(
        2,
        column_name_to_number("D"),
        "Часы переработки:",
        &format,
    )?;

    format = cell_style(DataType::UsualText, CellType::InputHeader);
    // Hours header
    month_worksheet.write_with_format(2, column_name_to_number("B"), "Часы", &format)?;
    // Salary input header
    month_worksheet.write_with_format(4, column_name_to_number("D"), "Оклад:", &format)?;

    format = cell_style(DataType::UsualText, CellType::TotalPayment);
    // Total payout header
    month_worksheet.write_with_format(5, column_name_to_number("D"), "К получению:", &format)?;

    // Month
    format = match first_day.season() {
        Season::Winter => cell_style(DataType::UsualText, CellType::MonthWinter),
        Season::Spring => cell_style(DataType::UsualText, CellType::MonthSpring),
        Season::Summer => cell_style(DataType::UsualText, CellType::MonthSummer),
        Season::Autumn => cell_style(DataType::UsualText, CellType::MonthAutumn),
    };

    month_worksheet.merge_range(
        1,
        column_name_to_number("A"),
        1,
        column_name_to_number("C"),
        first_day.month_name().as_str(),
        &format,
    )?;
    Ok(())
}
