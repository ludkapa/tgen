use anyhow::Result as AResult;
use rust_xlsxwriter::{Formula, utility::column_name_to_number, worksheet::Worksheet};

pub(crate) fn add_total_cells(
    month_worksheet: &mut Worksheet,
    work_hours: u16,
    total_days: u8,
    usual_days_formula: String,
    weekend_formula: String,
) -> AResult<()> {
    // Work hours
    month_worksheet.write(0, column_name_to_number("E"), work_hours)?;
    // Overtime hours formula
    month_worksheet.write_formula(
        1,
        column_name_to_number("E"),
        Formula::new(usual_days_formula),
    )?;
    // Weekend hours formula
    month_worksheet.write_formula(2, column_name_to_number("E"), Formula::new(weekend_formula))?;
    // Total payment formula
    month_worksheet.write_formula(
        5,
        column_name_to_number("E"),
        Formula::new(format!("=SUM(C4:C{})+E5", total_days + 4)),
    )?;
    Ok(())
}
