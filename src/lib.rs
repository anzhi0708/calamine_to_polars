use std::{error::Error, fmt::Display, fs::File, io::BufReader, path::Path};

use calamine::{CellType, Data, DataType, Error as CalamineError, Range, Reader, Xlsx};
use polars::prelude::*;
use polars::{frame::DataFrame, series::Series};

pub struct CalamineToPolarsReader {
    workbook: Xlsx<BufReader<File>>,
}

/// Implelemt pandas style type catsing API for specified column(s).
pub trait ConvenientCast<'a> {
    fn with_types(
        &mut self,
        col_and_type: &'a [(&'a str, polars::datatypes::DataType)],
    ) -> Result<DataFrame, Box<dyn Error>>;
}

impl<'a> ConvenientCast<'a> for DataFrame {
    fn with_types(
        &mut self,
        col_and_type: &'a [(&'a str, polars::datatypes::DataType)],
    ) -> Result<DataFrame, Box<dyn Error>> {
        let mut all_series: Vec<Series> = Vec::new();
        for column in self.get_columns() {
            let mut is_column_added = false;
            for (col_name, col_cast_type) in col_and_type {
                if col_name == column.name() {
                    all_series.push(column.cast(col_cast_type).unwrap());
                    is_column_added = true;
                }
            } // end of inner for
            if !is_column_added {
                all_series.push(column.to_owned());
            }
        } // end of outer for
        Ok(DataFrame::new(all_series).unwrap())
    }
}

/// Implement API interfaces on calamine Range<T>
/// to convert calamine Excel data to Polars DataFrame.
///
pub trait ToPolarsDataFrame {
    /// This method assumes the input calamine Excel data
    /// has headers (column titles).
    /// It tries to convert Excel data into strongly-typed DataFrame.
    fn to_frame_auto_type(&mut self) -> Result<DataFrame, Box<dyn Error>>;
    /// Convert to DataFrame but everything's a String
    fn to_frame_all_str(&self) -> Result<DataFrame, Box<dyn Error>>;
}

impl<T> ToPolarsDataFrame for Range<T>
where
    T: DataType + CellType + Display,
{
    fn to_frame_all_str(&self) -> Result<DataFrame, Box<dyn Error>> {
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
            .map(|(col, name)| Series::new((&name).into(), col))
            .collect();

        // constructing DataFrame
        let df = DataFrame::new(series)?;

        Ok(df)
    }

    fn to_frame_auto_type(&mut self) -> Result<DataFrame, Box<dyn Error>> {
        let mut columns: Vec<Series> = Vec::new();
        let mut column_types: Vec<polars::datatypes::DataType> = Vec::new();
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
            column_types.push(polars::datatypes::DataType::Null);
        }

        // The first row of the ramaining part decides each column's data type
        for (col_index, cell) in self.rows().nth(1).unwrap().iter().enumerate() {
            let header = headers[col_index].as_str();
            match cell {
                c if c.is_int() => {
                    column_types[col_index] = polars::datatypes::DataType::Int64;
                    columns.push(Series::new(header.into(), [cell.get_int().unwrap()]));
                }
                c if c.is_float() => {
                    column_types[col_index] = polars::datatypes::DataType::Float64;
                    columns.push(Series::new(header.into(), [cell.get_float().unwrap()]));
                }
                c if c.is_bool() => {
                    column_types[col_index] = polars::datatypes::DataType::Boolean;
                    columns.push(Series::new(header.into(), [cell.get_bool().unwrap()]));
                }
                c if c.is_string() => {
                    column_types[col_index] = polars::datatypes::DataType::String;
                    columns.push(Series::new(header.into(), [cell.get_string().unwrap()]));
                }
                c if c.is_empty() => {
                    column_types[col_index] = polars::datatypes::DataType::Null;
                    columns.push(Series::new(
                        header.into(),
                        [cell.get_string().unwrap_or_default()],
                    ));
                }
                c if c.is_error() => {
                    panic!("This cell is error. The first row of the ramaining part decides each column's data type");
                }
                _ => {
                    panic!("Unknown error. The first row of the ramaining part decides each column's data type");
                }
            }
            // todo!()
        }
        dbg!(DataFrame::new(columns.clone()).unwrap());

        // iterating through all rows remaining
        for (row_index, row) in self.rows().skip(2).enumerate() {
            for (col_idx, cell) in row.iter().enumerate() {
                let header = headers[col_idx].as_str();
                match cell {
                    c if c.is_int() => {
                        let new_series = Series::new(header.into(), [c.get_int()]);

                        let append_result = columns[col_idx].append(&new_series);
                        match append_result {
                            Ok(_) => {}
                            Err(_) => {
                                eprintln!(
                                    "{}",
                                    format!("row {row_index}, col {header} (column index {col_idx}): expected int").as_str()
                                );
                                dbg!(&new_series);
                            }
                        }
                        /*
                        columns[col_idx].append(&new_series).expect(
                                format!("row {row_index}, col {header} (column index {col_idx}): expected int").as_str()
                        );
                        */
                    }
                    c if c.is_float() => {
                        let new_series = Series::new(header.into(), [c.get_float()]);

                        let append_result = columns[col_idx].append(&new_series);
                        match append_result {
                            Ok(_) => {}
                            Err(_) => {
                                eprintln!(
                                    "{}",
                                    format!("row {row_index}, col {header} (column index {col_idx}): expected float").as_str()
                                );
                                dbg!(&new_series);
                            }
                        }
                        /*
                        columns[col_idx].append(&new_series).expect(
                                format!("row {row_index}, col {header} (column index {col_idx}): expected float").as_str()
                        );
                        */
                        /*
                        columns[col_idx].append(&new_series).expect(
                            format!("row {row_index}, col {header} (column index {col_idx}): expected float").as_str(),
                        );
                        */
                    }
                    c if c.is_bool() => {
                        let new_series = Series::new(header.into(), [c.get_bool()]);
                        columns[col_idx].append(&new_series).expect(
                            format!("row {row_index}, col {header} (column index {col_idx}): expected bool").as_str(),
                        );
                    }
                    c if c.is_string() => {
                        let new_series = Series::new(header.into(), [c.get_string()]);
                        columns[col_idx].append(&new_series).expect(
                            format!("row {row_index}, col {header} (column index {col_idx}): expected string").as_str(),
                        );
                    }
                    c if c.is_empty() => {
                        let new_series = Series::new_empty(
                            header.into(),
                            polars::datatypes::DataType::Null.as_ref(),
                        );
                        columns[col_idx].append(&new_series).unwrap();
                    }
                    _ => {
                        panic!("Error when reading all data...")
                    }
                }
            }
        }

        let df = DataFrame::new(columns)?;

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
