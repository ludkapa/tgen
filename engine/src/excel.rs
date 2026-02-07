use crate::excel::data::days::Day;
use crate::excel::data::days::DayType;
use crate::excel::data::days::Days;
use crate::excel::data::days::add_day_cell;
use crate::excel::data::headers::add_header_cells;
use crate::excel::data::total::add_total_cells;
use crate::excel::design::CellType;
use crate::excel::design::DataType;
use crate::excel::design::cell_style;

use anyhow::Ok;
use anyhow::Result as AResult;
use chrono::{Datelike, NaiveDate, Weekday};
use network::holiday;
use rust_xlsxwriter::utility::column_name_to_number;
use rust_xlsxwriter::workbook::Workbook;
use std::collections::HashSet;

pub async fn get_filled_table(year: u16, salary: u32) -> AResult<Vec<u8>> {
    // Creating table
    let mut table = Workbook::new();
    // Get all days in selected year
    let days: Days = get_dates_at_year(year).await?;
    // Split days to chunks by month
    let chunks = days.split_months();
    for month_days in chunks {
        // Creating Sheet
        let month_worksheet = table.add_worksheet();
        // For total formulas
        let mut weekend_formula: String = "=".to_string();
        let mut usual_days_formula: String = "=".to_string();
        // Adding headers
        add_header_cells(month_worksheet, month_days.first().unwrap())?;
        // For work hours
        let mut work_hours: u16 = 0;
        // Iterate over days in month chunk
        for day in month_days {
            // Adding day to month sheet and geting this flag
            let flag = add_day_cell(month_worksheet, day)?;
            // Creating formula for total block
            match flag {
                DayType::Usual => {
                    usual_days_formula = format!("{}B{}+", usual_days_formula, 3 + day.number());
                    work_hours = work_hours + 8;
                }
                _ => {
                    weekend_formula = format!("{}B{}+", weekend_formula, 3 + day.number());
                }
            }
        }
        // Remove last plus sign from formulas
        usual_days_formula.pop();
        weekend_formula.pop();
        // Add salary
        let salary = match salary {
            0 => "".to_string(),
            _ => salary.to_string(),
        };

        let format = cell_style(DataType::Money, CellType::InputHeader);
        month_worksheet.write_with_format(4, column_name_to_number("E"), salary, &format)?;
        // Add a total block
        add_total_cells(
            month_worksheet,
            work_hours,
            month_days.len() as u8,
            usual_days_formula,
            weekend_formula,
        )?;
        // Autofit columns
        month_worksheet.autofit();
        // Make E column wider
        month_worksheet.set_column_width(column_name_to_number("E"), 12)?;
        // Set month name
        month_worksheet.set_name(month_days.first().unwrap().month_name())?;
    }
    // Convert struct to bytes and return it
    let buf = table.save_to_buffer()?;
    Ok(buf)
}

async fn get_dates_at_year(year: u16) -> AResult<Days> {
    if year < 2017 {
        anyhow::bail!("You cant use year lowest than 2017!");
    }
    let raw_holidays = holiday::fetch_holidays_by_year(year).await?;
    let holidays: HashSet<NaiveDate> = raw_holidays.into_iter().collect();
    let first_date = NaiveDate::from_ymd_opt(year as i32, 1, 1).unwrap();

    let days: Days = first_date
        .iter_days()
        .take_while(|d| d.year() == year as i32)
        .map(|d| {
            if d.weekday() == Weekday::Sun {
                Day::new(d, DayType::Weekend)
            } else if d.weekday() == Weekday::Sat || holidays.contains(&d) {
                Day::new(d, DayType::Earn)
            } else {
                Day::new(d, DayType::Usual)
            }
        })
        .collect();

    Ok(days)
}

mod data;
mod design;
mod network;
