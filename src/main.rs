use std::path::Path;
use calamine::{Reader, open_workbook, Xlsx};

fn main() {
    let path = Path::new("./assets/items.xlsx");

    // opens a new workbook
    let mut workbook: Xlsx<_> = open_workbook(path).expect("Cannot open file");
    let sheets = workbook.worksheets();

    // check if sheets exists
    if sheets.is_empty() {
        return;
    }

    let range = &sheets[0].1;
    let start = range.start().unwrap();
    let end = range.end().unwrap();

    for row in start.0..=end.0 {
        for col in start.1..=end.1 {
            let pos = (row, col);
            println!("{:#?}", range.get_value(pos).unwrap());
        }
    }

    // println!("{:#?}", range.rows());
    println!("{:#?}, {:#?}", range.start(), range.end());

    // for (name, range) in sheets {
    //     println!("{:#?} | {:#?}", name, range);
    // }

    // Now get all formula!
    // let sheets = workbook.sheet_names().to_owned();
    // for s in sheets {
    //     println!("found {} formula in '{}'",
    //         workbook
    //         .worksheet_formula(&s)
    //         .expect("sheet not found")
    //         .expect("error while getting formula")
    //         .rows().flat_map(|r| r.iter().filter(|f| !f.is_empty()))
    //         .count(),
    //         s);
    // }
}
