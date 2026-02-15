use anyhow::Ok;
use anyhow::Result as AResult;
use chrono::NaiveDate;
use serde::Deserialize;

#[derive(Deserialize, Debug, PartialEq)]
pub(crate) struct FetchedDates {
    holidays: Vec<NaiveDate>,
}

pub async fn fetch_holidays_by_year(year: u16) -> AResult<Vec<NaiveDate>> {
    let url = "https://raw.githubusercontent.com/d10xa/holidays-calendar/refs/heads/master/json/calendar.json";
    let response = reqwest::get(url).await?.error_for_status()?;
    let dates_str = match response.json::<FetchedDates>().await.ok() {
        Some(dates) => dates,
        None => return Err(anyhow::anyhow!("Holidays not fetched!")),
    };
    let year_str = year.to_string();
    let holidays_at_year: Vec<NaiveDate> = dates_str
        .holidays
        .into_iter()
        .filter(|d| d.contains(&year_str))
        .filter_map(|d| NaiveDate::parse_from_str(d.as_str(), "%Y-%m-%d").ok())
        .collect();
    Ok(holidays_at_year)
}
