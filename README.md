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

    Multi-threading for fast process



🛠 In Development:

    Column extraction functionality


🛠 How to build (install cargo if you dont already have it):

    cargo install --release excelactor

Usage

    Run the executable(type excelactor in command line)

    Enter your search keyword

    Enter Sheet name (leave empty if you want all sheets searched)

    Specify the folder containing Excel files

    Results will be saved in results.xlsx
