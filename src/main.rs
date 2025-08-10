extern crate umya_spreadsheet;
use std::fs::{self, File};
use std::io::{BufRead, BufReader, Write};
use std::io;
use std::path::{Path, PathBuf};
use regex::Regex;
use umya_spreadsheet::Worksheet;
use walkdir::WalkDir;
use anyhow::{Context, Result};
use std::time::Duration;
use std::error::Error;
use std::sync::mpsc;
use std::thread;
use std::collections::BTreeMap;
use std::sync::{Arc, Mutex};
use umya_spreadsheet::*;
struct coordinates {
    row: u32,
    column: u32,
}

fn select_folder() -> Option<PathBuf> {
    rfd::FileDialog::new()
    .set_title("Select a folder containing XLSX files")
    .pick_folder()
}

fn find_xlsx_files(folder: &Path) -> Result<Vec<PathBuf>> {
    let mut xlsx_files = Vec::new();

    for entry in WalkDir::new(folder) {
        let entry = entry?;
        let path = entry.path();

        if path.is_file() {
            if let Some(ext) = path.extension() {
                if ext == "xlsx" {
                    xlsx_files.push(path.to_path_buf());
                }
            }
        }
    }

    Ok(xlsx_files)
}

//checks if our row is in the same range as merged cells
fn check_range(merged: &String, selected: &str) -> bool {
    let re = Regex::new(r"^[A-Za-z](\d{1,3}):[A-Za-z](\d{1,3})$").unwrap();
    let caps = match re.captures(merged) {
        Some(c) => c,
        None => {
            return false
        }
    };

    let num1 = caps.get(1)
    .unwrap_or_else(|| panic!("No group 1 in string '{}'", merged))
    .as_str()
    .parse::<u32>()
    .unwrap_or_else(|_| panic!("Failed to parse group 1 in string '{}'", merged));

    let num2 = caps.get(2)
    .unwrap_or_else(|| panic!("No group 2 in string '{}'", merged))
    .as_str()
    .parse::<u32>()
    .unwrap_or_else(|_| panic!("Failed to parse group 2 in string '{}'", merged));

    let selected_row = selected.parse().unwrap();
    if num1 < selected_row && selected_row < num2 {
        return true
    }
    else {
        return false
    }
}


fn get_row(row: u32, sheet: &Worksheet) -> Vec<String> {    
    let mut row_values = Vec::new();
    let merged = sheet.get_merge_cells();
    let cell_row = row.to_string();
    //for sorting merged rows
    let mut rowmap = BTreeMap::new();

    for range in merged {
       let mut range_value = range.get_range();
    if check_range(&range_value, &cell_row) == true {
        let mut merge_coord = sheet.map_merged_cell(&*range_value);
        let mut value = sheet.get_value(merge_coord);
        let column_num = merge_coord.0;
            rowmap.insert(column_num, value.to_string());
        
    }
   }

    let cell = sheet.get_collection_by_row(&row);
    for item in cell {
        let column = item.get_coordinate().get_col_num();
        let value = item.get_cell_value().get_value();
        rowmap.insert(*column, value.to_string());
    }

    for (key, val) in rowmap.range(0..){
            row_values.push(val.to_string());
    }
    row_values
}


//creates a vector of everything in the row
fn get_keyword_coord(query: &str, sheet: &Worksheet) -> Vec<coordinates>
{
    let mut coords = Vec::new();
    let cells = sheet.get_cell_collection();
    for item in cells {
        let mut value = item.get_cell_value().get_value();
        if query == value{
            coords.push(coordinates {
                row: *item.get_coordinate().get_row_num(),
                column: *item.get_coordinate().get_col_num(),
            });
        }
    }
    coords
}
fn prompt_input(prompt: &str) -> io::Result<String> {
    let mut input = String::new();
    print!("{}", prompt);
    io::stdout().flush()?; // Ensure prompt appears immediately
    io::stdin().read_line(&mut input)?;
    Ok(input.trim().to_string())
}
fn row_writer(row: Vec<String>, row_num: u32) {
    
    
    
}
fn main() {
    println!("Please select a folder containing the excel files");
    let folder = select_folder().ok_or(anyhow::anyhow!("No folder selected")).unwrap();
    let xlsx_files = find_xlsx_files(&folder).unwrap();
    let length = xlsx_files.len();
    println!("Found xlsx files");
    // Get the query
    let keyword = prompt_input("Enter your search query: ").expect("Failed to read query");

    // Get sheet name
    let sheet = prompt_input("Enter Sheet name: ").expect("Failed to read");
    let (tx, rx) = mpsc::channel();
    let mut handles = vec![];
    let mut counter = 0;
    let mut wtr = Writer::from_path("output.csv").unwrap();
    let counter = Arc::new(Mutex::new(0));
    
    for file in xlsx_files {

        let keyword = keyword.to_string();
        let sheet = sheet.clone();
        let tx = tx.clone();
        let counter = Arc::clone(&counter);
     let handle = thread::spawn(move || {
        let book = umya_spreadsheet::reader::xlsx::read(&file).unwrap();
        let sheet = book.get_sheet_by_name(&sheet).unwrap();
        let coords = get_keyword_coord(&keyword, &sheet);
        let filename = &file.file_name().unwrap().to_str().unwrap();
        let mut results = vec![];
        
        for cord in coords {
            let mut row = get_row(cord.row, &sheet);
            if row.len() != 0 {
                row.insert(0, filename.to_string()); // Add filename as first column
                results.push(row);
            }

        }
        tx.send(results).unwrap();
        let mut num = counter.lock().unwrap();
        *num += 1;
        print!("\rProcessed {}/{} files", *num, length);
        io::stdout().flush().unwrap();
    });
    
     handles.push(handle);

    }
    drop(tx);
    // Collect results
     let mut flush_counter = 0;
     let mut results = new_file();
     let result_sheet = results.new_sheet("RESULTS").unwrap();

     let mut row_ind = 1;
     let mut column_ind = 1;
     for received in rx {
         let rows  = received;
         for row in rows {
                for str in &row {
                results.get_sheet_mut(&1).unwrap().get_cell_mut((&column_ind, &row_ind)).set_value(str);
                column_ind += 1;
             }
            column_ind = 1;
            row_ind += 1;
         }
     }
     
     for handle in handles {
        handle.join().unwrap();
    }
    let path = std::path::Path::new("./results.xlsx");
    let _ = writer::xlsx::write(&results, path);
    println!("/nProcess finished");    
      let output_path = "output.csv";

    println!("Successfully produced results. Output written to {}", output_path);  

}
