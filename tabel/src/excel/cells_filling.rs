use anyhow::Result as AResult;
use rust_xlsxwriter::{
    FormatBorder, Formula, utility::column_name_to_number, worksheet::Worksheet,
};

use crate::{
    entities::days::{Day, DayType, Season},
    excel::styles::{CellType, DataType, cell_style},
};

pub(super) fn add_day_cell(month_worksheet: &mut Worksheet, day: &Day) -> AResult<DayType> {
    let day_row = 2 + day.number();

    let mut format = match day.earn_type() {
        DayType::Earn => cell_style(DataType::UsualText, CellType::Earn),
        DayType::Weekend => cell_style(DataType::UsualText, CellType::Weekend),
        DayType::Usual => cell_style(DataType::UsualText, CellType::Usual),
    };
    month_worksheet.write_with_format(day_row, column_name_to_number("B"), "0", &format)?;
    month_worksheet.write_with_format(
        day_row,
        0,
        format!("{} {}", day.number(), day.weekday_short()),
        &format.set_border_left(FormatBorder::Medium),
    )?;

    format = cell_style(DataType::Money, CellType::TotalBonus);
    month_worksheet.write_formula_with_format(
        day_row,
        column_name_to_number("C"),
        Formula::new(format!("=E5/E1*B{}*2", day_row + 1)),
        &format,
    )?; //Complite it
    Ok(day.earn_type())
}

pub(super) fn add_header_cells(month_worksheet: &mut Worksheet, first_day: &Day) -> AResult<()> {
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

pub(super) fn add_formulas_cells(
    month_worksheet: &mut Worksheet,
    work_hours: u16,
    total_days: u8,
    usual_days_formula: String,
    weekend_formula: String,
) -> AResult<()> {
    let mut format = cell_style(DataType::UsualText, CellType::Header);
    // Work hours
    month_worksheet.write_with_format(0, column_name_to_number("E"), work_hours, &format)?;
    // Overtime hours formula
    month_worksheet.write_formula_with_format(
        2,
        column_name_to_number("E"),
        Formula::new(usual_days_formula),
        &format,
    )?;
    // Weekend hours formula
    month_worksheet.write_formula_with_format(
        1,
        column_name_to_number("E"),
        Formula::new(weekend_formula),
        &format,
    )?;

    format = cell_style(DataType::Money, CellType::TotalPayment);
    // Total payment formula
    month_worksheet.write_formula_with_format(
        5,
        column_name_to_number("E"),
        Formula::new(format!("=SUM(C4:C{})+E5", total_days + 4)),
        &format,
    )?;
    Ok(())
}
