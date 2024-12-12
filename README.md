# json-excel-matcher-with-nodejs

This project compares a JSON format file with a specific column of an Excel format file and saves the matched and unmatched data in separate Excel files. The project is particularly focused on normalizing complex characters in the data and increasing the similarity rate with fuzzy matching.

## 🚀 Features

- Completely scans the JSON file and converts it to a flat list.
- Avoids confusion by normalizing special characters in JSON and Excel data.
- Converts data from an Excel file to JSON format and normalizes the data in a given column.
- It increases the similarity rate and detects matches using fuzzy matching.
- Saves matched and unmatched data separately in Excel files.

## 🛠️ Requirements

- **Node.js**
- Necessary Node.js libraries such as xlsx and fs (to be installed in the project).
  ```bash
   npm install xlsx
   npm install fuse.js

## 🔧 Installation

1. Clone or download this project:
   ```bash
   git clone https://github.com/KenancanA/json-excel-matcher-with-nodejs
   cd json-excel-matcher-with-nodejs
2. Install the required dependencies:
   ```bash
   npm install 
4. Add the following files to the project directory:
   - data.json: Your data file in JSON format.
   - table.xlsx: Your data file in Excel format.

## 🔧 Usage

1. Start the comparison by running match.js:
   ```bash
   node match.js
3. By default, the code compares the “server name” column in the Excel file with the data in the JSON file. If you want to compare another column, edit the corresponding code in the match.js file to specify the column name you want to compare.
