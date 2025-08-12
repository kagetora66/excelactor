Excelactor

A Rust tool for bulk searching and extracting data from multiple Excel files

Overview

This tool  processes a folder of Excel files (.xlsx), searches for rows containing a specified keyword, and exports the results to an excel file..
Key Features

✔ Bulk processing - Scan multiple Excel files in one operation

✔ Keyword search - Extract rows containing your search term

✔ Fast Performance: Rust's concurrency features are fully utilized for max performance

✔ Merged cell support - Properly handles merged rows/columns

✔ Output is exported as an excel file

Current Status

✅ Implemented:

    Row extraction with keyword matching

    Merged row support



🛠 In Development:

    Column extraction functionality


🛠 How to build:

    cargo build --release

Usage

    Run the executable

    Enter your search keyword

    Enter Sheet name

    Specify the folder containing Excel files

    Results will be saved in output.csv
