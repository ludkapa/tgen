use crate::excel::design::{CellStyle, CellType, DataType};
use anyhow::Result as AResult;
use chrono::{Datelike, NaiveDate, Weekday};
use derive_more::{Deref, DerefMut, IntoIterator};
use rust_xlsxwriter::{Formula, utility::column_name_to_number, worksheet::Worksheet};

#[derive(Default, Debug, Clone, Copy)]
pub(crate) enum DayType {
    #[default]
    Usual,
    Earn,
    Weekend,
}

pub(crate) enum Season {
    Winter,
    Spring,
    Summer,
    Autumn,
}

#[derive(Default, Debug)]
pub(crate) struct Day {
    day: NaiveDate,
    flag: DayType,
}

impl Day {
    pub(crate) fn new(day: NaiveDate, flag: DayType) -> Self {
        Self { day, flag }
    }

    pub(crate) fn year(&self) -> i32 {
        self.day.year()
    }

    pub(crate) fn number(&self) -> u32 {
        self.day.day()
    }

    pub(crate) fn weekday_short(&self) -> String {
        match self.day.weekday() {
            Weekday::Mon => "Пн".to_string(),
            Weekday::Tue => "Вт".to_string(),
            Weekday::Wed => "Ср".to_string(),
            Weekday::Thu => "Чт".to_string(),
            Weekday::Fri => "Пт".to_string(),
            Weekday::Sat => "Сб".to_string(),
            Weekday::Sun => "Вс".to_string(),
        }
    }

    pub(crate) fn month_name(&self) -> String {
        match self.day.month() {
            1 => "❄️ Январь".to_string(),
            2 => "🌨️ Февраль".to_string(),
            3 => "🌱 Март".to_string(),
            4 => "🌸 Апрель".to_string(),
            5 => "🌿 Май".to_string(),
            6 => "☀️ Июнь".to_string(),
            7 => "🏖️ Июль".to_string(),
            8 => "🍉 Август".to_string(),
            9 => "🍂 Сентябрь".to_string(),
            10 => "🍁 Октябрь".to_string(),
            11 => "🌧️ Ноябрь".to_string(),
            12 => "🎄 Декабрь".to_string(),
            _ => "❓ Неизвестный месяц".to_string(),
        }
    }

    pub(crate) fn season(&self) -> Season {
        match self.day.month() {
            1 | 2 | 12 => Season::Winter,
            3 | 4 | 5 => Season::Spring,
            6 | 7 | 8 => Season::Summer,
            9 | 10 | 11 => Season::Autumn,
            _ => Season::Winter,
        }
    }
}

#[derive(IntoIterator, Deref, DerefMut)]
pub(crate) struct Days(Vec<Day>);

impl FromIterator<Day> for Days {
    fn from_iter<T: IntoIterator<Item = Day>>(iter: T) -> Self {
        Self(iter.into_iter().collect())
    }
}

impl Days {
    pub(crate) fn split_months(&self) -> impl Iterator<Item = &[Day]> {
        self.chunk_by(|a, b| a.day.month() == b.day.month())
    }
}

pub(crate) fn add_day_cell(month_worksheet: &mut Worksheet, day: &Day) -> AResult<DayType> {
    let day_row = 2 + day.number();

    let mut format = match day.flag {
        DayType::Earn => CellStyle(DataType::UsualText, CellType::Earn)?,
        DayType::Weekend => CellStyle(DataType::UsualText, CellType::Weekend)?,
        DayType::Usual => CellStyle(DataType::UsualText, CellType::Usual)?,
    };
    month_worksheet.write_with_format(
        day_row,
        0,
        format!("{} {}", day.number(), day.weekday_short()),
        &format,
    )?;
    month_worksheet.write_with_format(day_row, column_name_to_number("B"), "0", &format)?;

    format = CellStyle(DataType::Money, CellType::TotalBonus)?;
    month_worksheet.write_formula_with_format(
        day_row,
        column_name_to_number("C"),
        Formula::new(format!("=E5/E1*B{}", day_row + 1)),
        &format,
    )?; //Complite it
    Ok(day.flag)
}
