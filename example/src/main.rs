use calamine_to_polars::*;
use core::error::Error;
use polars::datatypes::DataType::{Float32, Int32};
use std::process::exit;

fn main() {
    println!("Hello, world!");
    test_df();
}

fn test_df() -> Result<(), Box<dyn Error>> {
    use std::env::args;
    if args().count() < 2 {
        eprintln!("None of the required 2 command-line arguments `file path`, `sheet name` was not provided");
        exit(-1);
    }
    if args().count() < 3 {
        eprintln!("The required command-line argument `sheet name` was not provided");
        exit(-1);
    }

    let file_path = args().nth(1).unwrap();
    let sheet_name = args().nth(2).unwrap();

    let mut df = CalamineToPolarsReader::new(file_path)
        .open_sheet(sheet_name)
        .unwrap()
        .to_frame_all_str()?;

    // Before convenient casting
    println!("{:#?}", df);
    println!("{:#?}", df["수량"]);

    df = df.with_types(&[
        // Change column name to match yours
        ("상품합계", Float32),
        // Change column name to match yours
        ("수량", Int32),
    ])?;

    // After convenient casting
    println!("{:#?}", df["수량"]);
    println!("{:#?}", df);


    let gb = df.group_by(["Group"])?;
    let groups = gb.groups();

    for element in groups.iter() {
        println!("{:?}", element);
    }

    let parts = df.partition_by(["Group"], false)?;
    for part in parts {
        println!("{:#?}", part);
    }
    
    Ok(())
}
