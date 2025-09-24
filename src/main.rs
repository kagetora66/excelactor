extern crate umya_spreadsheet;
use std::io::{BufRead,Write};
use std::io;
use std::path::{Path, PathBuf};
use regex::Regex;
use umya_spreadsheet::Worksheet;
use walkdir::WalkDir;
use anyhow::{Context, Result};
use std::sync::mpsc;
use std::thread;
use std::collections::BTreeMap;
use std::sync::{Arc, Mutex};
use umya_spreadsheet::*;
struct coordinates {
    row: u32,
    column: u32,
}
#[derive(Debug, Clone, PartialEq)]
enum ExtractState {
    ExtractRow,
    ExtractColumn,
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
fn check_range(merged: &String, selected: &str, State: ExtractState) -> bool {
    let re = Regex::new(r"^(A-Za-z)(\d{1,3}):(A-Za-z)(\d{1,3})$").unwrap();
    let caps = match re.captures(merged) {
        Some(c) => c,
        None => {
            return false
        }
    };
    if let ExtractState::ExtractRow = State {
    let num1 = caps[2].parse::<u32>().unwrap_or(0);
    let num2 = caps[4].parse::<u32>().unwrap_or(0);
    let selected_row = selected.parse().unwrap();
    num1 < selected_row && selected_row < num2
   }
    else {
    let num1 = caps[1].chars()
        .map(|c| c.to_ascii_uppercase() as u32 - 'A' as u32 + 1)
        .fold(0, |acc, digit| acc * 26 + digit);
    let num2 = caps[3].chars()
        .map(|c| c.to_ascii_uppercase() as u32 - 'A' as u32 + 1)
        .fold(0, |acc, digit| acc * 26 + digit);
    let selected_column = selected.parse().unwrap();
    num1 < selected_column && selected_column < num2
    }
}

fn get_row(row: u32, sheet: &Worksheet) -> Vec<String> {    
    let mut row_values = Vec::new();
    let merged = sheet.get_merge_cells();
    let cell_row = row.to_string();
    //for sorting merged rows
    let mut rowmap = BTreeMap::new();
    for range in merged {
       let range_value = range.get_range();
    if check_range(&range_value, &cell_row, ExtractState::ExtractRow) == true {
        let merge_coord = sheet.map_merged_cell(&*range_value);
        let value = sheet.get_value(merge_coord);
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
fn get_column(column: u32, sheet: &Worksheet) -> Vec<String> {    
    let mut column_values = Vec::new();
    let merged = sheet.get_merge_cells();
    let cell_column = column.to_string();
    //for sorting merged column
    let mut columnmap = BTreeMap::new();
    for range in merged {
       let range_value = range.get_range();
    if check_range(&range_value, &cell_column, ExtractState::ExtractRow) == true {
        let merge_coord = sheet.map_merged_cell(&*range_value);
        let value = sheet.get_value(merge_coord);
        let column_num = merge_coord.0;
            columnmap.insert(column_num, value.to_string());
        
    }
   }

    let cell = sheet.get_collection_by_column(&column);
    for item in cell {
        let row = item.get_coordinate().get_row_num();
        let value = item.get_cell_value().get_value();
        columnmap.insert(*row, value.to_string());
    }

    for (key, val) in columnmap.range(0..){
            column_values.push(val.to_string());
    }
    column_values
}

//creates a vector of everything in the row or column
fn get_keyword_coord(query: &str, sheet: &Worksheet) -> Vec<coordinates>
{
    let mut coords = Vec::new();
    let cells = sheet.get_cell_collection();
    let mut Query = String::new();
    Query = query.to_lowercase();
    for item in cells {
        let value = item.get_cell_value().get_value();
        let Value = value.to_lowercase();
        if Value.contains(&Query){
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
fn row_writer(rows: Vec<Vec<String>>, sheet: &mut Spreadsheet) {
    
    let mut column_ind = 1;
    let mut row_ind = 1;
    for row in rows {
        for str in &row {
            sheet.get_sheet_mut(&0).unwrap().get_cell_mut((&column_ind, &row_ind)).set_value(str);
            column_ind += 1;
            }
        column_ind = 1;
        row_ind += 1;
        }
}
fn column_writer(columns: Vec<Vec<String>>, sheet: &mut Spreadsheet) {
    
    let mut column_ind = 1;
    let mut row_ind = 1;
    for column in columns {
        for str in &column {
            sheet.get_sheet_mut(&0).unwrap().get_cell_mut((&column_ind, &row_ind)).set_value(str);
            row_ind += 1;
            }
        row_ind = 1;
        column_ind += 1;
        }
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
    let sheet = prompt_input("Enter Sheet name (leave empty if you want all sheets searched): ").expect("Failed to read");
    //Get the state of extraction (column or rows)
    let input_state = prompt_input("Do you want rows or columns containing the keyword to be extracted? (enter r or c)").expect("Failed to read");
    let extract_state = if input_state.trim().eq_ignore_ascii_case("c") {
    ExtractState::ExtractColumn
    } else if input_state.trim().eq_ignore_ascii_case("r") {
    ExtractState::ExtractRow
    } else {
    // Default or error handling
    println!("Invalid input defaulting to extract columns");
    ExtractState::ExtractColumn
    };
    let (tx, rx) = mpsc::channel();
    let mut handles = vec![];
    let counter = Arc::new(Mutex::new(0));
    
    for file in xlsx_files {

        let keyword = keyword.to_string();
        let sheet = sheet.clone();
        let tx = tx.clone();
        let counter = Arc::clone(&counter);
        let extract_state = extract_state.clone();
     let handle = thread::spawn(move || {
        let book = umya_spreadsheet::reader::xlsx::read(&file).unwrap();
        let sheet_list: &[Worksheet];
        //Search through all sheets or just one
        if sheet == ""{ 
            sheet_list = book.get_sheet_collection();
        }
        else{
            let sheet_ref = book.get_sheet_by_name(&sheet).unwrap();
            sheet_list = std::slice::from_ref(sheet_ref);
        }
       
        let filename = &file.file_name().unwrap().to_str().unwrap();
        let mut results = vec![];
        if let ExtractState::ExtractRow = extract_state {
            for sheet in sheet_list{
	     let coords = get_keyword_coord(&keyword, sheet);
	     for cord in coords {
              let mut row = get_row(cord.row, &sheet);
              if row.len() != 0 {
                 row.insert(0, filename.to_string()); // Add filename as first column
                 row.insert(1, sheet.get_name().to_string()); // Add sheet as second column
                 results.push(row);
                }
            }
	    }
        }
         else{
	     for sheet in sheet_list{
	     let coords = get_keyword_coord(&keyword, sheet);
		 for cord in coords {
		     let mut column = get_column(cord.column, &sheet);
                     if column.len() != 0 {
			 column.insert(0, filename.to_string()); // Add filename as first row
			 column.insert(1, sheet.get_name().to_string()); // Add sheet as second row
			 results.push(column);
		     }
		 }
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
     let mut results = new_file();
     let mut all_results = Vec::new();
     for received in rx {
         all_results.extend(received);
     }
     all_results.sort();
     all_results.dedup();
     if let ExtractState::ExtractRow = extract_state {
         row_writer(all_results, &mut results);
        }
        else {
         column_writer(all_results, &mut results);

        }
     
     for handle in handles {
        handle.join().unwrap();
    }
    let path = std::path::Path::new("./results.xlsx");
    let _ = writer::xlsx::write(&results, path);
    println!("\nProcess finished");    
      let output_path = "results.xlsx";

    println!("Successfully produced results. Output written to {}", output_path);
}
