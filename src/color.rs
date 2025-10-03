//We first store all coordinates that indicate the beginning and end of files then get and color them
use umya_spreadsheet::{self, Color};
use std::path::Path;
fn find_file_coordinates(results_xls: &Path) -> Vec<String>{
   let mut coordinates = vec![];
   let book = umya_spreadsheet::reader::xlsx::read(&results_xls).expect("File not found");
   let sheet = book.get_sheet_by_name("Sheet1").unwrap();
   let cells = sheet.get_cell_collection();
   for item in cells {
       let content = item.get_cell_value().get_value();
       if content == "File Name" || content == "Sheet Name" {
          coordinates.push(item.get_coordinate().get_coordinate());
       }
       if content.contains("End of"){
           coordinates.push(item.get_coordinate().get_coordinate());
       }
           
    }
   return coordinates
   
}
pub fn colour_separators(results_xls: &Path) -> () {

    let coloring_cells = find_file_coordinates(results_xls);
    let mut book = umya_spreadsheet::reader::xlsx::read(&results_xls).expect("File not found");
    let mut sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();
    let mut cells = sheet.get_cell_collection_mut();
    for item in cells {
        for coord in coloring_cells.clone(){
            if item.get_coordinate().get_coordinate() == coord {
                let mut style = item.get_style_mut();
                style.set_background_color(Color::COLOR_DARKBLUE);
            }
        }
    }
    umya_spreadsheet::writer::xlsx::write(&book, results_xls).expect("Failed to write file");
}
