use anyhow::Ok;
use anyhow::Result as AResult;
use chrono::Datelike;
use chrono::NaiveDate;
use serde::Deserialize;

#[derive(Deserialize, Debug, PartialEq)]
pub(crate) struct FetchedDates {
    holidays: Vec<NaiveDate>,
    #[serde(skip)]
    exist_years: Option<Vec<u16>>,
}

impl FetchedDates {
    pub(crate) async fn init() -> AResult<Self> {
        let url = "https://raw.githubusercontent.com/d10xa/holidays-calendar/refs/heads/master/json/calendar.json";
        let response = reqwest::get(url).await?.error_for_status()?;
        let fetched_dates = match response.json::<FetchedDates>().await.ok() {
            Some(dates) => dates,
            None => return Err(anyhow::anyhow!("Cannot fetch holidays from github!")),
        };
        Ok(fetched_dates)
    }

    pub(crate) fn get_year_dates(&self, year: u16) -> Vec<NaiveDate> {
        self.holidays
            .iter()
            .copied()
            .filter(|d| d.year() == year as i32)
            .collect()
    }

    pub(crate) fn get_exist_years(&mut self) -> Vec<u16> {
        if let Some(cached) = &self.exist_years {
            return cached.clone();
        }
        let mut years: Vec<u16> = self.holidays.iter().map(|d| d.year() as u16).collect();
        years.sort();
        years.dedup();
        self.exist_years = Some(years.clone());
        years
    }
}

// pub async fn fetch_holidays_by_year(year: u16) -> AResult<Vec<NaiveDate>> {
//     let year_str = year.to_string();
//     let holidays_at_year: Vec<NaiveDate> = dates_str
//         .holidays
//         .into_iter()
//         .filter(|d| d.contains(&year_str))
//         .filter_map(|d| NaiveDate::parse_from_str(d.as_str(), "%Y-%m-%d").ok())
//         .collect();
//     Ok(holidays_at_year)
// }
