# Calamine -> Polars DataFrame

## Example

```rust
use calamine_to_polars::*;
use polars::frame::DataFrame;

fn main() -> Result<(), calamine::Error> {

    // Loading Excel
    let mut reader = CalamineToPolarsReader::new("YOUR_DATA.xlsx");

    // Use `to_df()` to get a Polars DataFrame:
    let df: DataFrame = reader.open_sheet("Sheet1").unwrap().to_df().unwrap();

    // println!("{:#?}", df); // Prints the DataFrame


    // Get Sheet1 column titles
    let column_names = reader.get_column_names("Sheet1")?;
    println!("{:#?}", column_names);


    Ok(())
}
```
