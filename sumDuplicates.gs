function sumDuplicatesInSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  // Get all data
  const data = sheet.getDataRange().getValues();

  // Object to store sums by unique key in column A
  const sums = {};

  for (let i = 0; i < data.length; i++) {
    const key = data[i][0]; // Column A
    const value = parseFloat(data[i][1]) || 0; // Column B, force number

    if (sums[key] === undefined) {
      sums[key] = value;
    } else {
      sums[key] += value;
    }
  }

  // Clear existing sheet
  sheet.clear();

  // Write back unique entries with sums
  const output = Object.entries(sums);
  sheet.getRange(1, 1, output.length, 2).setValues(output);
}
