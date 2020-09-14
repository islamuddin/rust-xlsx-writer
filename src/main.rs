extern crate umya_spreadsheet;

fn main() {
    
    // reader
    let path = std::path::Path::new("F:/rust/coders/xls_sample/aaa.xlsx");
    //let book = umya_spreadsheet::reader::xlsx::read(path).unwrap();
    // or
    // new file
    let mut book = umya_spreadsheet::new_file();
    
    // new worksheet
    let _ = book.new_sheet("Sheet2");
    
    // change value
    let _ = book.get_sheet_by_name_mut("Sheet2").unwrap().get_cell_mut("A1").set_value("TEST1");
    // or
    let _ = book.get_sheet_mut(1).get_cell_by_column_and_row_mut(1, 1).set_value("TEST1");
    
    // read value
    let a1_value = book.get_sheet_by_name("Sheet2").unwrap().get_cell("A1").unwrap().get_value();
    // or
    let a1_value = book.get_sheet(1).unwrap().get_cell_by_column_and_row(1, 1).unwrap().get_value();
    assert_eq!("TEST1", a1_value);  // TEST1
    
    // add bottom border
    let _ = book.get_sheet_by_name_mut("Sheet2").unwrap()
    .get_style_mut("A1")
    .get_borders_mut()
    .get_bottom_mut()
    .set_border_style(umya_spreadsheet::Border::BORDER_MEDIUM);
    
    // writer
    let path = std::path::Path::new("F:/rust/coders/xls_sample/aaa.xlsx");
    let _ = umya_spreadsheet::writer::xlsx::write(&book, path);

}
