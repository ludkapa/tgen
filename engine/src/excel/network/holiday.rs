use std::collections::HashSet;

use anyhow::Result as AResult;
use chrono::Datelike;
use chrono::NaiveDate;
use serde::Deserialize;

#[derive(Deserialize, Debug, PartialEq)]
pub struct FetchedDates {
    holidays: HashSet<NaiveDate>,
}

impl FetchedDates {
    pub async fn init() -> AResult<Self> {
        let url = "https://raw.githubusercontent.com/d10xa/holidays-calendar/refs/heads/master/json/calendar.json";
        let response = reqwest::get(url).await?.error_for_status()?;
        let mut fetched_dates = response.json::<FetchedDates>().await?;
        let last_year = match fetched_dates.holidays.iter().map(|d| d.year()).max() {
            Some(year) => year,
            None => return Err(anyhow::anyhow!("Cannot get last year!")),
        };
        fetched_dates.holidays = fetched_dates
            .holidays
            .into_iter()
            .filter(|d| d.year() == last_year)
            .collect();
        Ok(fetched_dates)
    }

    pub(crate) fn get_holidays(&self) -> HashSet<NaiveDate> {
        self.holidays.clone()
    }
}
