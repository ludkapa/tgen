use anyhow::Result as AResult;
use rust_xlsxwriter::{Formula, utility::column_name_to_number, worksheet::Worksheet};

use crate::excel::design::{CellType, DataType, cell_style};

pub(crate) fn add_total_cells(
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
        1,
        column_name_to_number("E"),
        Formula::new(usual_days_formula),
        &format,
    )?;
    // Weekend hours formula
    month_worksheet.write_formula_with_format(
        2,
        column_name_to_number("E"),
        Formula::new(weekend_formula),
        &format,
    )?;

    format = cell_style(DataType::Money, CellType::Earn);
    // Total payment formula
    month_worksheet.write_formula_with_format(
        5,
        column_name_to_number("E"),
        Formula::new(format!("=SUM(C4:C{})+E5", total_days + 4)),
        &format,
    )?;
    Ok(())
}
