function exportSpreadsheetSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  const schema = {
    spreadsheetName: ss.getName(),
    generatedAt: new Date().toISOString(),
    sheets: []
  };

  sheets.forEach(sheet => {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    let headers = [];

    if (lastRow > 0 && lastCol > 0) {
      headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    }

    schema.sheets.push({
      sheetName: sheet.getName(),
      totalRows: lastRow,
      totalColumns: lastCol,
      columns: headers.map((header, index) => ({
        position: index + 1,
        name: header || `Column_${index + 1}`
      }))
    });
  });

  const file = DriveApp.createFile(
    "CRM_Schema.json",
    JSON.stringify(schema, null, 2),
    MimeType.PLAIN_TEXT
  );

  Logger.log("Schema file created: " + file.getUrl());
}