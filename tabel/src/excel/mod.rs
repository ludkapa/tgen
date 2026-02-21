use crate::{
    adapter::holidays::FetchedDates,
    entities::days::{DayType, Days},
    excel::{
        cells_filling::{add_day_cell, add_formulas_cells, add_header_cells},
        styles::{CellType, DataType, cell_style},
    },
};
use anyhow::Result as AResult;
use rust_xlsxwriter::{
    Format, FormatBorder, Worksheet, utility::column_name_to_number, workbook::Workbook,
};

pub async fn get_filled_table(salary: u32) -> AResult<Vec<u8>> {
    // Creating table
    let mut table = Workbook::new();
    // Fetch holidays
    let holidays = FetchedDates::init().await?;
    // Generate days for filling
    let days = Days::new_with_holidays(holidays.get_holidays());
    // Split days to chunks by month
    let chunks = days.split_months();
    for month_days in chunks {
        // Creating Sheet
        let month_worksheet = table.add_worksheet();
        // Make worksheet white
        month_worksheet.set_screen_gridlines(false);
        // For total formulas
        let mut weekend_cells = Vec::new();
        let mut usual_day_cells = Vec::new();
        // For work hours
        let mut work_hours: u16 = 0;
        // Iterate over days in month chunk
        for day in month_days {
            // Row number
            let row_number = 3 + day.number();
            // Adding day to month sheet and geting this flag
            let flag = add_day_cell(month_worksheet, day)?;
            // Creating formula for total block
            match flag {
                DayType::Usual => {
                    usual_day_cells.push(format!("B{}", row_number));
                    work_hours = work_hours + 8;
                }
                _ => {
                    weekend_cells.push(format!("B{}", row_number));
                }
            }
        }

        // Adding headers
        add_header_cells(month_worksheet, month_days.first().unwrap())?;

        // Make formula string
        let weekends_formula = format!("={}", weekend_cells.join("+"));
        let usual_days_formula = format!("={}", usual_day_cells.join("+"));
        // Add salary
        let salary_str = match salary {
            0 => "".to_string(),
            _ => salary.to_string(),
        };

        let mut format = cell_style(DataType::Money, CellType::InputHeader);
        month_worksheet.write_formula_with_format(
            4,
            column_name_to_number("E"),
            format!("={}", salary_str).as_str(),
            &format,
        )?;
        // Add a total block
        add_total_cells(
            month_worksheet,
            work_hours,
            month_days.len() as u8,
            usual_days_formula,
            weekends_formula,
        )?;
        // Polish worksheet
        // Do wider border at bottom of days block
        format = Format::new().set_border_top(FormatBorder::Medium);
        month_worksheet.merge_range(
            3 + month_days.len() as u32,
            column_name_to_number("A"),
            3 + month_days.len() as u32,
            column_name_to_number("C"),
            "",
            &format,
        )?;
        // Autofit columns
        month_worksheet.autofit();
        // Make E column wider
        month_worksheet.set_column_width(column_name_to_number("E"), 12)?;
        // Make C column wider
        month_worksheet.set_column_width(column_name_to_number("C"), 10)?;
        // Make B column narrower
        month_worksheet.set_column_width(column_name_to_number("B"), 7.5)?;
        // Set month name
        month_worksheet.set_name(month_days.first().unwrap().month_name())?;
    }
    // Convert struct to bytes and return it
    let buf = table.save_to_buffer()?;
    Ok(buf)
}

mod cells_filling;
mod styles;
