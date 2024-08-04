use calamine::{CellType, Data, DataType, Error as CalamineError, Range, Reader, Xlsx};
use polars::prelude::*;
use polars::{frame::DataFrame, series::Series};
use std::{error::Error, fmt::Display, fs::File, io::BufReader, path::Path};

pub struct CalamineToPolarsReader {
    workbook: Xlsx<BufReader<File>>,
}

pub trait ToPolarsDataFrame {
    fn to_df(&mut self) -> Result<DataFrame, Box<dyn Error>>;
}

impl<T> ToPolarsDataFrame for Range<T>
where
    T: DataType + CellType + Display,
{
    fn to_df(&mut self) -> Result<DataFrame, Box<dyn Error>> {
        let mut columns = Vec::new();

        // Headers
        let headers: Vec<String> = self
            .rows()
            .next()
            .ok_or("No data")?
            .iter()
            .map(|cell| cell.to_string())
            .collect();

        // Vec<String> for each column
        for _ in 0..headers.len() {
            columns.push(Vec::<String>::new());
        }

        // iterating through all rows
        for row in self.rows().skip(1) {
            for (col_idx, cell) in row.iter().enumerate() {
                columns[col_idx].push(cell.to_string());
            }
        }

        // list of `Series`s
        let series: Vec<Series> = columns
            .into_iter()
            .zip(headers)
            .map(|(col, name)| Series::new(&name, col))
            .collect();

        // constructing DataFrame
        let df = DataFrame::new(series)?;

        Ok(df)
    }
}

impl CalamineToPolarsReader {
    //
    pub fn open_workbook<P: AsRef<Path>>(file_name: P) -> Xlsx<BufReader<File>> {
        let workbook: Xlsx<_> =
            calamine::open_workbook(file_name).expect("Could not open workbook");
        workbook
    }

    pub fn new<P: AsRef<Path>>(file_name: P) -> Self {
        Self {
            workbook: CalamineToPolarsReader::open_workbook(file_name),
        }
    }

    //
    pub fn open_sheet<S: AsRef<str>>(&mut self, sheet_name: S) -> Option<Range<Data>> {
        if let Ok(sheet_range) = self.workbook.worksheet_range(sheet_name.as_ref()) {
            Some(sheet_range)
        } else {
            None
        }
    }

    //
    pub fn get_column_names<S: AsRef<str>>(
        &mut self,
        sheet_name: S,
    ) -> Result<Vec<String>, CalamineError> {
        if let Ok(sheet_range) = self.workbook.worksheet_range(sheet_name.as_ref()) {
            let width = sheet_range.width();

            let mut column_names = Vec::<String>::new();
            for idx in 0..width {
                let _column_name = sheet_range.get_value((0u32, idx as u32)).unwrap();
                let column_name: String = format!("{}", _column_name);
                column_names.push(column_name);
            }
            return Ok(column_names);
        }

        return Err(CalamineError::Msg("Missing column name"));
    }
}
