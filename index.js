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
    debugger;
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
      cell.forEach((item) => {
        let value = item.text;
        if (Number(value) || Number(value) === 0) value = Number(value);
        sheet.getCell(rowNo, columnNo).value = value;
      });
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

const extractPivotData = (canvas) => {
  let columnWidth = 0;
  let columnAxisLength = [];

  // Exporting data from the column matrix
  const columnMatrix = canvas._composition.layout._columnMatrix._tree.matrix;
  let columnHeaders = new Array(columnMatrix.length);

  columnMatrix.forEach((row, i) => {
    let prev = 0;
    let prevHeader = "";
    columnHeaders[i] = row.flatMap((cell, j) => {
      if (cell._source === null) {
        return "";
      } else if (typeof cell.source() === "object") {
        const axisCellLength = cell._source._domain.length;
        columnWidth += axisCellLength;
        columnAxisLength.push(axisCellLength);

        if (i > 0 && typeof columnMatrix[i - 1][j]._source === "string") {
          let k = i;

          while (k > 0) {
            columnHeaders[k - 1].splice(
              j + 1 + prev,
              0,
              ...Array(axisCellLength - 1).fill("")
            );
            k -= 1;
          }
          prev += axisCellLength - 1;
        }
        return cell._source._domain;
      } else {
        if (prevHeader === cell._source) {
          return "";
        } else {
          prevHeader = cell._source;
          return prevHeader;
        }
      }
    });
  });
  // console.log("columnHeaders", columnHeaders);

  // Exporting data from the row matrix
  const rowMatrix = canvas._composition.layout._rowMatrix._tree.matrix;
  const rowMatLength = rowMatrix.length;
  const columnLength = rowMatrix[0].length;

  let rowHeaders = [];
  let rowWidth = 0;
  let rowAxisLength = [];

  rowMatrixIter: for (let i = columnMatrix.length; i < rowMatLength; i++) {
    let row = [];
    for (let j = 0; j < columnLength; j++) {
      const cell = rowMatrix[i][j];
      if (cell._source === null) {
        row.push("");
      } else if (typeof cell._source === "string") {
        row.push(cell._source);
      } else {
        const domain = cell._source._domain.reverse();
        const axisCellLength = domain.length;
        rowWidth += axisCellLength;
        rowAxisLength.push(axisCellLength);

        row.push(domain[0]);
        if (axisCellLength > 1) {
          rowHeaders.push(row);
          for (let k = 1; k < axisCellLength; k++) {
            let extraRow = Array(columnLength - 1).fill("");
            extraRow.push(domain[k]);
            rowHeaders.push(extraRow);
            continue rowMatrixIter;
          }
        }
      }
    }
    rowHeaders.push(row);
  }
  // console.log("rowHeaders", rowHeaders);
  console.log("columnWidth", columnWidth);

  console.log("rowWidth", rowWidth);
  console.log("columnAxisLength", columnAxisLength);
  console.log("rowAxisLength", rowAxisLength);

  // Exporting data from the geom matrix
  const geomMatrix = canvas._composition.layout._centerMatrix._layoutMatrix;
  //   console.log(canvas._composition.layout._centerMatrix);
  let geomData = Array.from({ length: rowWidth }, () =>
    Array(columnWidth).fill(null)
  );

  let prevY = 0;
  for (let i = 0; i < geomMatrix[0].length; i++) {
    let prevX = 0;
    for (let j = 0; j < geomMatrix.length; j++) {
      let data = geomMatrix[j][i]._source._layers[0]._pointMap;

      for (let k = 0; k < Object.keys(data).length; k++) {
        const dataPoint = data[k];

        if (
          dataPoint.yIndex + prevX < rowWidth &&
          dataPoint.xIndex + prevY < columnWidth
        ) {
          if (
            geomData[dataPoint.yIndex + prevX][dataPoint.xIndex + prevY] ===
            null
          ) {
            geomData[dataPoint.yIndex + prevX][dataPoint.xIndex + prevY] = [];
          }
          //   console.log(canvas._composition.layout._centerMatrix);
          geomData[dataPoint.yIndex + prevX][dataPoint.xIndex + prevY].push(
            dataPoint
          );
        }
      }
      prevX += rowAxisLength[j];
    }
    prevY += columnAxisLength[i];
  }
  console.log("geomData", geomData);

  const xSplit = rowMatrix[0].length;
  const ySplit = columnMatrix.length;
  return {
    columnHeaders: columnHeaders,
    rowHeaders: rowHeaders,
    xSplit: xSplit,
    ySplit: ySplit,
    geomData: geomData,
  };
};

export const exportToExcel = (canvas) => {
  const data = extractPivotData(canvas);
  convertToExcel(data);
};
