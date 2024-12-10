const xlsx = require("xlsx");
const fs = require("fs");
const Fuse = require("fuse.js");

// 1. Read JSON File
const jsonData = JSON.parse(fs.readFileSync("data.json", "utf-8"));

// Scan the complete JSON and convert all values into a flat list
const flattenJSONValues = (obj) => {
  const values = [];
  const recursiveFlatten = (current) => {
    if (typeof current === "object" && current !== null) {
      for (const key in current) {
        recursiveFlatten(current[key]);
      }
    } else if (typeof current === "string" || typeof current === "number") {
      values.push(normalize(current.toString())); // Add as normalized
    }
  };
  recursiveFlatten(obj);
  return values;
};

// 2. Normalize Function (Uppercase-Lowercase, Turkish Character and Special Character Correction)
const normalize = (value) =>
  value
    .toLowerCase()
    .trim()
    .replace(/[\W_]+/g, "") // Remove special characters
    .replace(/ç/g, "c")
    .replace(/ğ/g, "g")
    .replace(/ş/g, "s")
    .replace(/ı/g, "i")
    .replace(/ö/g, "o")
    .replace(/ü/g, "u");

// Output all values in JSON in normalized form
const jsonValues = flattenJSONValues(jsonData);
console.log("JSON'dan Normalize Edilen Değerler (Toplam):", jsonValues.length);

// 3. Read Excel File
const workbook = xlsx.readFile("tablo.xlsx");
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];

// Convert Excel data to JSON format
const excelData = xlsx.utils.sheet_to_json(sheet);
console.log("Excel'den Alınan Toplam Satır Sayısı:", excelData.length);

// Normalize the data in the “Server Name” column of Excel
const excelServerNames = excelData.map((row) => normalize(row["Sunucu Adı"]));
console.log(
  "Excel'den Normalize Edilen Sunucu Adları (Toplam):",
  excelServerNames.length
);

// 4. Fuzzy Matching Settings
const fuse = new Fuse(jsonValues, {
  includeScore: true,
  threshold: 0.2, // 80% similarity
});

// Find a close match and save matches
const matches = excelData.filter((row) => {
  const normalizedValue = normalize(row["Sunucu Adı"]);
  const result = fuse.search(normalizedValue);
  return result.length > 0;
});

// Find mismatches
const nonMatches = excelData.filter((row) => {
  const normalizedValue = normalize(row["Sunucu Adı"]);
  const result = fuse.search(normalizedValue);
  return result.length === 0;
});

// Check matches and non-matches
console.log("Eşleşen Sunucu Sayısı:", matches.length);
console.log("Eşleşmeyen Sunucu Sayısı:", nonMatches.length);

// Check: Matched + Non-Matched = Total Number of Rows
if (matches.length + nonMatches.length !== excelData.length) {
  console.error("HATA: Toplam satır sayısı uyuşmuyor!");
  console.log("Eşleşen + Eşleşmeyen:", matches.length + nonMatches.length);
  console.log("Excel Toplam Satır Sayısı:", excelData.length);
} else {
  console.log("Eşleşme kontrolü doğru.");
}

// 5. Save Results to New Excel Files (Server Name Only)
if (matches.length > 0) {
  // Print only the “Server Name” column
  const matchesSheetData = matches.map((row) => ({
    "Sunucu Adı": row["Sunucu Adı"],
  }));
  const matchesSheet = xlsx.utils.json_to_sheet(matchesSheetData);
  const matchesWorkbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(matchesWorkbook, matchesSheet, "Matches");
  xlsx.writeFile(matchesWorkbook, "eslesenler.xlsx");
  console.log("Eşleşen veriler 'eslesenler.xlsx' dosyasına kaydedildi.");
}

if (nonMatches.length > 0) {
  // Print only the “Server Name” column
  const nonMatchesSheetData = nonMatches.map((row) => ({
    "Sunucu Adı": row["Sunucu Adı"],
  }));
  const nonMatchesSheet = xlsx.utils.json_to_sheet(nonMatchesSheetData);
  const nonMatchesWorkbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(
    nonMatchesWorkbook,
    nonMatchesSheet,
    "NonMatches"
  );
  xlsx.writeFile(nonMatchesWorkbook, "eslesmeyenler.xlsx");
  console.log("Eşleşmeyen veriler 'eslesmeyenler.xlsx' dosyasına kaydedildi.");
} else {
  console.log("Eşleşmeyen veri bulunamadı.");
}
