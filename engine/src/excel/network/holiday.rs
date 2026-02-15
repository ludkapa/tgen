use anyhow::Result as AResult;
use chrono::Datelike;
use chrono::NaiveDate;
use serde::Deserialize;

#[derive(Deserialize, Debug, PartialEq)]
pub struct FetchedDates {
    holidays: Vec<NaiveDate>,
    #[serde(skip)]
    exist_years: Option<Vec<u16>>,
}

impl FetchedDates {
    pub async fn init() -> AResult<Self> {
        let url = "https://raw.githubusercontent.com/d10xa/holidays-calendar/refs/heads/master/json/calendar.json";
        let response = reqwest::get(url).await?.error_for_status()?;
        let fetched_dates = match response.json::<FetchedDates>().await.ok() {
            Some(dates) => dates,
            None => return Err(anyhow::anyhow!("Cannot fetch holidays from github!")),
        };
        Ok(fetched_dates)
    }

    pub(crate) fn for_year(&self, year: u16) -> Vec<NaiveDate> {
        self.holidays
            .iter()
            .copied()
            .filter(|d| d.year() == year as i32)
            .collect()
    }

    pub fn available_years(&mut self) -> Vec<u16> {
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
