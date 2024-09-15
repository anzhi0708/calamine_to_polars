use calamine_to_polars::*;
use polars::datatypes::DataType::{Float32, Int32};
use std::process::exit;

fn main() {
    println!("Hello, world!");
    test_df();
}

fn test_df() {
    use std::env::args;
    if args().count() < 2 {
        eprintln!("No file path, no sheet name found");
        exit(-1);
    }
    if args().count() < 3 {
        eprintln!("No sheet name found");
        exit(-1);
    }

    let file_path = args().nth(1).unwrap();
    let sheet_name = args().nth(2).unwrap();

    let mut df = CalamineToPolarsReader::new(file_path)
        .open_sheet(sheet_name)
        .unwrap()
        .to_frame_all_str()
        .unwrap();

    // Before convenient casting
    println!("{:#?}", df);

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
}
