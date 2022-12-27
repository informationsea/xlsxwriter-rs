#[cfg(feature = "chrono")]
use chrono::{Datelike, Timelike};

use super::DateTime;

impl DateTime {
    pub fn new(year: i16, month: i8, day: i8, hour: i8, min: i8, second: f64) -> DateTime {
        DateTime {
            year,
            month,
            day,
            hour,
            min,
            second,
        }
    }

    pub fn date(year: i16, month: i8, day: i8) -> DateTime {
        DateTime {
            year,
            month,
            day,
            hour: 0,
            min: 0,
            second: 0.0,
        }
    }

    pub fn time(hour: i8, min: i8, second: f64) -> DateTime {
        DateTime {
            year: 0,
            month: 0,
            day: 0,
            hour,
            min,
            second,
        }
    }
}

#[cfg(feature = "chrono")]
impl From<chrono::naive::NaiveDateTime> for DateTime {
    fn from(datetime: chrono::naive::NaiveDateTime) -> Self {
        DateTime {
            year: datetime.year() as i16,
            month: datetime.month() as i8,
            day: datetime.day() as i8,
            hour: datetime.hour() as i8,
            min: datetime.minute() as i8,
            second: datetime.second().into(),
        }
    }
}

#[cfg(feature = "chrono")]
impl From<chrono::naive::NaiveDate> for DateTime {
    fn from(datetime: chrono::naive::NaiveDate) -> Self {
        DateTime {
            year: datetime.year() as i16,
            month: datetime.month() as i8,
            day: datetime.day() as i8,
            hour: 0,
            min: 0,
            second: 0.,
        }
    }
}

#[cfg(feature = "chrono")]
impl From<chrono::naive::NaiveTime> for DateTime {
    fn from(datetime: chrono::naive::NaiveTime) -> Self {
        DateTime {
            year: 1900,
            month: 1,
            day: 0,
            hour: datetime.hour() as i8,
            min: datetime.minute() as i8,
            second: datetime.second().into(),
        }
    }
}

impl From<&DateTime> for libxlsxwriter_sys::lxw_datetime {
    fn from(datetime: &DateTime) -> Self {
        libxlsxwriter_sys::lxw_datetime {
            year: datetime.year.into(),
            month: datetime.month.into(),
            day: datetime.day.into(),
            hour: datetime.hour.into(),
            min: datetime.min.into(),
            sec: datetime.second,
        }
    }
}
