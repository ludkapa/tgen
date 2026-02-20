use crate::{
    entities::days::{Day, DayType},
    excel::styles::{CellType, DataType, cell_style},
};
use anyhow::Result as AResult;
use rust_xlsxwriter::{
    FormatBorder, Formula, utility::column_name_to_number, worksheet::Worksheet,
};

pub(crate) fn add_day_cell(month_worksheet: &mut Worksheet, day: &Day) -> AResult<DayType> {
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
