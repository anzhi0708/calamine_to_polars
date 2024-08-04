# Calamine -> Polars DataFrame

## Example

```rust
use calamine_to_polars::*;

fn main() -> Result<(), calamine::Error> {

    let mut reader = CalamineToPolarsReader::new("your_data.xlsx");

    // Specify the sheet's name, e.g. "Sheet1"
    let df = reader.open_sheet("Sheet1").unwrap().to_df();

    println!("{:#?}", df);

    // Getting header names
    let column_names = reader.get_column_names("Sheet1")?;

    Ok(())
}
```
