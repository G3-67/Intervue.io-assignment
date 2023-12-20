const fs = require('fs');
const XLSX = require('xlsx');

function convertNestedJSONToExcel(json, outputFileName) {
  const workbook = XLSX.utils.book_new();

  // Sheet 1: Basic Information with Test Columns
  const sheet1Data = [
    {
      "Name": json.name,
      "Location": json.location,
      "Is Open": json.isOpen,
      "Number of Sections": json.numberOfSections,
      "Contact": json.contact,
      "Popular Genres": json.popularGenres.join(', '), // Join genres if it's an array
    }
  ];
  const sheet1 = XLSX.utils.json_to_sheet(sheet1Data, { skipHeader: true });
  XLSX.utils.book_append_sheet(workbook, sheet1, 'Sheet1');

  // Sheet 2: Empty Sheet
  const sheet2 = XLSX.utils.aoa_to_sheet([]);
  XLSX.utils.book_append_sheet(workbook, sheet2, 'Sheet2');

  // Sheet 3: Test Object
  const sheet3Data = [];
  if (json.test) {
    for (const key in json.test) {
      const row = { "Test Key": key, "Test Value": json.test[key] };
      sheet3Data.push(row);
    }
  }
  const sheet3 = XLSX.utils.json_to_sheet(sheet3Data, { skipHeader: true });
  XLSX.utils.book_append_sheet(workbook, sheet3, 'Sheet3');

  // Sheet 4: Section 1 with Title, Author, Price, Is Available Columns
  const sheet4Data = [];
  if (json.sections && json.sections[0] && json.sections[0].books) {
    json.sections[0].books.forEach((book, index) => {
      const row = { "Title": book.title, "Author": book.author, "Price": book.price, "Is Available": book.isAvailable, "Section": "Section 1" };
      sheet4Data.push(row);
    });
  }
  const sheet4 = XLSX.utils.json_to_sheet(sheet4Data, { skipHeader: true });
  XLSX.utils.book_append_sheet(workbook, sheet4, 'Sheet4');

  // Sheet 5: Section 2 with Title, Author, Price, Is Available Columns
  const sheet5Data = [];
  if (json.sections && json.sections[1] && json.sections[1].books) {
    json.sections[1].books.forEach((book, index) => {
      const row = { "Title": book.title, "Author": book.author, "Price": book.price, "Is Available": book.isAvailable, "Section": "Section 2" };
      sheet5Data.push(row);
    });
  }
  const sheet5 = XLSX.utils.json_to_sheet(sheet5Data, { skipHeader: true });
  XLSX.utils.book_append_sheet(workbook, sheet5, 'Sheet5');

  // Write the workbook to an Excel file
  XLSX.writeFile(workbook, outputFileName);
}

// Example nested JSON data
const nestedJSON = {
    "name": "The Reading Nook",
    "location": "123 Book St, Bibliopolis",
    "isOpen": true,
    "numberOfSections": 2,
    "contact": null,
    "popularGenres": ["Fiction", "Mystery", "Sci-Fi", "Non-Fiction"],
    "test": {
    "test1": "Test 1",
    "test2": {
    "test3": "Test 3"
    }
    },
    "sections": [
    {
    "sectionName": "Section 1",
    "books": [
    {
    "title": "Journey to the Unknown",
    "author": "Alice Wonder",
    "price": 12.99,
    "isAvailable": true
    },
    {
    "title": "Mystery of the Ancient Map",
    "author": "Clive Cussler",
    "price": 15.50,
    "isAvailable": false
    }
    ]
    },
    {
    "sectionName": "Section 2",
    "books": [
    {
    "title": "The Reality of Myths",
    "author": "Helen Troy",
    "price": 18.25,
    "isAvailable": true
    }
    ]
    }
    ]
}
        

// Specify the output file name
const outputFileName = 'output.xlsx';

// Convert nested JSON to Excel
convertNestedJSONToExcel(nestedJSON, outputFileName);

console.log(`Conversion successful. Check ${outputFileName}`);