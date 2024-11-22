import ExcelJS from "exceljs";

async function saveExcelFile(blob) {
  try {
    // Prompt user to choose a file location and name
    const fileHandle = await window.showSaveFilePicker({
      suggestedName: "newSheet.xlsx",
      types: [
        {
          description: "Excel Files",
          accept: {
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
              [".xlsx"],
          },
        },
      ],
    });

    // Create a writable stream to the file
    const writableStream = await fileHandle.createWritable();

    // Write the Blob content (Excel data) to the file
    await writableStream.write(blob);

    // Close the stream to save the file
    await writableStream.close();
  } catch (error) {
    console.error("Error saving file");
  }
}

const insertColumnHeaders = (columnHeaders, sheet) => {
  columnHeaders.forEach((row) => {
    sheet.addRow(row);
  });

  columnHeaders.forEach((rowData, rowIdx) => {
    let startIdx = 0;
    while (startIdx < rowData.length && rowData[startIdx] === "") startIdx++;

    if (startIdx > 0) sheet.mergeCells(rowIdx + 1, 1, rowIdx + 1, startIdx); //Need to pass the rowNo and columnNo
    if (startIdx === rowData.length) return;

    let endIdx = startIdx;
    while (startIdx < rowData.length) {
      if (rowData[startIdx] !== "" && startIdx !== endIdx) {
        sheet.mergeCells(rowIdx + 1, endIdx + 1, rowIdx + 1, startIdx);
        endIdx = startIdx;
      }
      startIdx++;
    }
    sheet.mergeCells(rowIdx + 1, endIdx + 1, rowIdx + 1, startIdx);
  });
};

const findLastNonEmptyRow = (data) => {
  for (let row = data.length - 1; row >= 0; row--) {
    // Check if all cells in the row are empty
    const isRowEmpty = data[row].every((cell) => cell === "");

    if (!isRowEmpty) {
      // 1 based indexing
      return row + 1;
    }
  }
  // If all rows are empty, return 0
  return 0;
};
const insertRowHeaders = (rowHeaders, sheet, prevRowCount) => {
  rowHeaders.forEach((row) => {
    sheet.addRow(row);
  });

  const numOfCols = rowHeaders[0].length;

  const rowLen = findLastNonEmptyRow(rowHeaders);
  // Traverse columns
  for (let col = 1; col <= numOfCols; col++) {
    let startRow = null;
    let mergeValue = "";

    for (let row = 1; row <= rowLen; row++) {
      const cellValue = rowHeaders[row - 1][col - 1];

      if (cellValue && cellValue !== mergeValue) {
        // If a new mergeValue is found, close the previous merge (if any)
        if (startRow !== null) {
          sheet.mergeCells(
            prevRowCount + startRow,
            col,
            prevRowCount + row - 1,
            col
          );

          sheet.getCell(prevRowCount + startRow, col).alignment = {
            vertical: "top",
          };
        }

        // Update mergeValue and startRow
        mergeValue = cellValue;
        startRow = row;
      } else if (!cellValue) {
        // If an empty cell is found, continue the merge
        if (startRow === null) {
          startRow = row;
        }
      }
    }

    // Merge the last range in the column with top alignment
    if (startRow !== null) {
      sheet.mergeCells(
        prevRowCount + startRow,
        col,
        prevRowCount + rowLen,
        col
      );

      sheet.getCell(prevRowCount + startRow, col).alignment = {
        vertical: "top",
      };
    }
  }
};

const insertGeomMatrix = (geomData, sheet, prevRowCount, prevColCount) => {
  let rowNo = prevRowCount + 1;
  const columnNoStart = prevColCount + 1;

  geomData.forEach((row) => {
    let columnNo = columnNoStart;

    row.forEach((cell) => {
      if (cell === null) {
        // Skip the cell if it's null
        columnNo++;
        return;
      }
      let cellValue = "";
      let align = "left";
      cell.forEach((item) => {
        let value = item.text;
        if (Number(value) || Number(value) === 0) {
          align = "right";
        }
        if (cellValue) cellValue += " ";
        cellValue += value;
      });
      // Insert the concatenated value into the worksheet cell
      const cellRef = sheet.getCell(rowNo, columnNo);
      cellRef.value = Number(cellValue) ? Number(cellValue) : cellValue;

      // Set wrapText to true for this cell
      cellRef.alignment = {
        wrapText: true,
        horizontal: align,
        vertical: "top",
      };
      columnNo++;
    });
    rowNo++;
  });
};
const convertToExcel = (params) => {
  const { columnHeaders, rowHeaders, xSplit, ySplit, geomData } = params;
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Sheet", {
    views: [
      {
        state: "frozen", // Enables frozen view
        xSplit: xSplit, // Freezes columns
        ySplit: ySplit, // Freezes rows
      },
    ],
  });

  insertColumnHeaders(columnHeaders, sheet);

  insertRowHeaders(rowHeaders, sheet, columnHeaders.length);

  insertGeomMatrix(
    geomData,
    sheet,
    columnHeaders.length,
    rowHeaders.length > 0 ? rowHeaders[0].length : 0
  );

  //Create a BLOB object from the workbook
  workbook.xlsx.writeBuffer().then((buffer) => {
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    saveExcelFile(blob);
  });
};

export const exportToExcel = (canvas) => {
  const data = extractPivotData(canvas);
  convertToExcel(data);
};
