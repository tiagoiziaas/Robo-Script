function main(workbook: ExcelScript.Workbook) {
  // Set the file path
  const filePath = "C:\Users\tiago\botao\EmpregadosemExcel.xls";

  // Open the file
  const file = Excel.run(filePath, (workbook) => {
    // Set the range where you want to remove empty rows and columns
    const range = workbook.getActiveWorksheet().getRange("A1:Z100");

    // Remove empty rows
    removeEmptyRows(range);

    // Remove empty columns
    removeEmptyColumns(range);

    // Save and close the file
    workbook.save();
    workbook.close();

    return workbook;
  });
}

function removeEmptyRows(range: ExcelScript.Range) {
  const rowCount = range.getRowCount();
  const columnCount = range.getColumnCount();

  // Loop through each row in the range
  for (let i = rowCount - 1; i >= 0; i--) {
    const row = range.getRow(i);
    let isEmpty = true;

    // Loop through each cell in the row
    for (let j = 0; j < columnCount; j++) {
      const cell = row.getCell(j);
      if (cell.getValue()!== "") {
        isEmpty = false;
        break;
      }
    }

    // If the row is empty, delete it
    if (isEmpty) {
      row.delete();
    }
  }
}

function removeEmptyColumns(range: ExcelScript.Range) {
  const rowCount = range.getRowCount();
  const columnCount = range.getColumnCount();

  // Loop through each column in the range
  for (let i = columnCount - 1; i >= 0; i--) {
    const column = range.getColumn(i);
    let isEmpty = true;

    // Loop through each cell in the column
    for (let j = 0; j < rowCount; j++) {
      const cell = column.getCell(j);
      if (cell.getValue()!== "") {
        isEmpty = false;
        break;
      }
    }

    // If the column is empty, delete it
    if (isEmpty) {
      column.delete();
    }
  }
}