use std::collections::HashMap;
use std::{error::Error, fmt::Display, fs::File, io::BufReader, path::Path};

use calamine::{CellType, Data, DataType, Error as CalamineError, Range, Reader, Xlsx};
use polars::frame::DataFrame;
use polars::prelude::*; // enum Column

pub struct CalamineToPolarsReader {
    workbook: Xlsx<BufReader<File>>,
}

/// Implelemt pandas style type catsing API for specified column(s).
pub trait CastColumnType<'a> {
    fn with_types(
        &mut self,
        col_and_type: &'a [(&'a str, polars::datatypes::DataType)],
    ) -> Result<DataFrame, Box<dyn Error>>;
}

impl<'a> CastColumnType<'a> for DataFrame {
    fn with_types(
        &mut self,
        col_and_type: &'a [(&'a str, polars::datatypes::DataType)],
    ) -> Result<DataFrame, Box<dyn Error>> {
        let mut all_columns: Vec<Column> = Vec::new();
        for column in self.get_columns() {
            let mut is_column_added = false;
            for (col_name, col_cast_type) in col_and_type {
                if col_name == column.name() {
                    all_columns.push(column.cast(col_cast_type).unwrap());
                    is_column_added = true;
                }
            } // end of inner for
            if !is_column_added {
                all_columns.push(column.to_owned());
            }
        } // end of outer for
        Ok(DataFrame::new(all_columns).unwrap())
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
    /// Pre-defined dtype(s) for upcoming DataFrame
    fn to_frame_with_types(&self, column_dtype: &HashMap<&str, polars::datatypes::DataType>);
}

impl<T> ToPolarsDataFrame for Range<T>
where
    T: DataType + CellType + Display,
{
    fn to_frame_with_types(&self, _column_dtype: &HashMap<&str, polars::datatypes::DataType>) {
        todo!();
    }

    fn to_frame_all_str(&self) -> Result<DataFrame, Box<dyn Error>> {
        let all_rows = self.rows().collect::<Vec<_>>();

        // iterating or duplicate headers
        let mut header_counts = HashMap::<String, usize>::new();
        let headers: Vec<String> = all_rows
            .first()
            .ok_or("No data")?
            .iter()
            .map(|cell| {
                let count = header_counts.entry(cell.to_string()).or_insert(0);
                let name = if *count > 0 {
                    format!("{}_duplicated_{}", cell, count)
                } else {
                    cell.to_string()
                };
                *count += 1;
                name
            })
            .collect();

        // pre allocated column memory
        let mut columns: Vec<Vec<String>> = vec![vec![]; headers.len()];
        columns.iter_mut().for_each(|v| v.reserve(all_rows.len()));

        // iterating through all rows
        for row in &all_rows[1..] {
            row.iter().enumerate().for_each(|(col_idx, cell)| {
                let cell_str = match cell {
                    c if c.is_datetime() => c
                        .as_datetime()
                        .map(|dt| dt.to_string())
                        .unwrap_or_else(|| String::new()),
                    _ => cell.to_string(),
                };
                columns[col_idx].push(cell_str);
            });
        }

        // list of `Column`s
        let columns: Vec<Column> = columns
            .into_iter()
            .zip(headers)
            .map(|(col, name)| Column::new((&name).into(), col))
            .collect();

        // constructing DataFrame
        let df = DataFrame::new(columns)?;

        Ok(df)
    }

    fn to_frame_auto_type(&mut self) -> Result<DataFrame, Box<dyn Error>> {
        let mut columns: Vec<Column> = Vec::new();
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
                    columns.push(Column::new(header.into(), [cell.get_int().unwrap()]));
                }
                c if c.is_float() => {
                    column_types[col_index] = polars::datatypes::DataType::Float64;
                    columns.push(Column::new(header.into(), [cell.get_float().unwrap()]));
                }
                c if c.is_bool() => {
                    column_types[col_index] = polars::datatypes::DataType::Boolean;
                    columns.push(Column::new(header.into(), [cell.get_bool().unwrap()]));
                }
                c if c.is_string() => {
                    column_types[col_index] = polars::datatypes::DataType::String;
                    columns.push(Column::new(header.into(), [cell.get_string().unwrap()]));
                }
                c if c.is_empty() => {
                    column_types[col_index] = polars::datatypes::DataType::Null;
                    columns.push(Column::new(
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
                        let new_column = Column::new(header.into(), [c.get_int()]);

                        let append_result = columns[col_idx].append(&new_column);
                        match append_result {
                            Ok(_) => {}
                            Err(_) => {
                                eprintln!(
                                    "{}",
                                    format!("row {row_index}, col {header} (column index {col_idx}): expected int").as_str()
                                );
                                dbg!(&new_column);
                            }
                        }
                        /*
                        columns[col_idx].append(&new_series).expect(
                                format!("row {row_index}, col {header} (column index {col_idx}): expected int").as_str()
                        );
                        */
                    }
                    c if c.is_float() => {
                        let new_column = Column::new(header.into(), [c.get_float()]);

                        let append_result = columns[col_idx].append(&new_column);
                        match append_result {
                            Ok(_) => {}
                            Err(_) => {
                                eprintln!(
                                    "{}",
                                    format!("row {row_index}, col {header} (column index {col_idx}): expected float").as_str()
                                );
                                dbg!(&new_column);
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
                        let new_column = Column::new(header.into(), [c.get_bool()]);
                        columns[col_idx].append(&new_column).expect(
                            format!("row {row_index}, col {header} (column index {col_idx}): expected bool").as_str(),
                        );
                    }
                    c if c.is_string() => {
                        let new_column = Column::new(header.into(), [c.get_string()]);
                        columns[col_idx].append(&new_column).expect(
                            format!("row {row_index}, col {header} (column index {col_idx}): expected string").as_str(),
                        );
                    }
                    c if c.is_empty() => {
                        let new_column = Column::new_empty(
                            header.into(),
                            polars::datatypes::DataType::Null.as_ref(),
                        );
                        columns[col_idx].append(&new_column).unwrap();
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
