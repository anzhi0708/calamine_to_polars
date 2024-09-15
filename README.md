# Calamine -> Polars DataFrame
⚠️ Under construction

## Example

```rust
use calamine_to_polars::*;
use polars::frame::DataFrame;

fn main() -> Result<(), Box<dyn Error>> {

    // Loading Excel
    //
    let file_path = "/path/to/your/excel.xlsx";
    let sheet_name = "sheet name";

    let mut df: DataFrame = CalamineToPolarsReader::new(file_path)
        .open_sheet(sheet_name)
        .unwrap()
        .to_frame_all_str()  // This method reads each cell's data as a string, you can cast a column's datatype later
        .unwrap();

    // Before casting
    println!("{:#?}", df);

    // Convenient cast
    df = df
        .with_types(&[
            // Change column name to match yours
            ("상품합계", Float32),
            // Change column name to match yours
            ("수량", Int32),
        ])
        .unwrap();

    // After convenient casting
    println!("{:#?}", df);


    Ok(())
}
```
