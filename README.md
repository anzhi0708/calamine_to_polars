# Calamine -> Polars DataFrame
⚠️ Under construction

## Example

```rust
use std::error::Error;

use calamine_to_polars::*;
use polars::frame::DataFrame;
use polars::datatypes::DataType::{Float32, Int32};


fn main() -> Result<(), Box<dyn Error>> {

    // Read Excel to DataFrame
    let file_path = "/path/to/your/excel.xlsx";
    // Sheet name, e.g. "Sheet1"
    let sheet_name = "Sheet1";
    let mut df: DataFrame = CalamineToPolarsReader::new(file_path)
        .open_sheet(sheet_name)
        .unwrap()
        .to_frame_all_str()  // This method reads each cell's data as a string, you can cast to some data type later
        .unwrap();

    println!("{:#?}", df);

    Ok(())
}
```
You can **specify the data type** for a specific column like this:
```rust
    // Specify the data type for a specific column using the `.with_types` method.
    df = df.with_types(&[
            // Change column name to match yours
            ("Qty", Int32),
            // Change column name to match yours
            ("Price", Float32),
        ]).unwrap();

    // Now column "Qty" is of type Int32, while "Price" is of type Float32.
    println!("{:#?}", df);
}
```
