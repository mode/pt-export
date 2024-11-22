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

const extractPivotData = (canvas) => {
  let columnWidth = 0;
  let columnAxisLength = [];
  // Exporting data from the column matrix
  const columnMatrix = canvas._composition.layout._columnMatrix._tree.matrix;
  let columnHeaders = new Array(columnMatrix.length);

  columnMatrix.forEach((row, i) => {
    let prev = 0;
    let prevHeader = "";
    headerMatrix[i] = row.flatMap((cell, j) => {
      if (cell._source === null) {
        return { text: "", style: {} };
      } else if (typeof cell.source() === "object") {
        const axisCellLength = cell._source._domain.length;
        columnWidth[0] += axisCellLength;
        columnAxisLength.push(axisCellLength);

        if (i > 0 && typeof matrix[i - 1][j]._source === "string") {
          let k = i;

          while (k > 0) {
            headerMatrix[k - 1].splice(
              j + 1 + prev,
              0,
              ...Array(axisCellLength - 1).fill({ text: "", style: {} })
            );
            k -= 1;
          }
          prev += axisCellLength - 1;
        }
        return cell._source._domain.map((element) => ({
          text: element,
          style: {},
        }));
      } else {
        if (prevHeader === cell._source) {
          return { text: "", style: {} };
        } else {
          prevHeader = cell._source;
          return { text: prevHeader, style: {} };
        }
      }
    });
  });

  return headerMatrix;
};

const generateLRMatrix = (matrix, headerMatrix, rowData) => {
  let { rowWidth, rowAxisLength } = rowData;
  // console.log("columnHeaders", columnHeaders);

  // Exporting data from the row matrix
  const rowMatrix = canvas._composition.layout._rowMatrix._tree.matrix;
  const rowMatLength = rowMatrix.length;
  const columnLength = rowMatrix[0].length;

  let rowHeaders = [];
  let rowWidth = 0;
  let rowAxisLength = [];

  rowMatrixIter: for (let i = 0; i < matrix.length; i++) {
    let row = [];
    for (let j = 0; j < matrix[0].length; j++) {
      const cell = matrix[i][j];
      if (cell._source === null) {
        row.push({ text: "", style: {} });
      } else if (typeof cell._source === "string") {
        row.push({ text: cell._source, style: {} });
      } else {
        const domain = cell._source._domain.reverse();
        const axisCellLength = domain.length;
        rowWidth[0] += axisCellLength;
        rowAxisLength.push(axisCellLength);

        row.push({ text: domain[0], style: {} });
        if (axisCellLength > 1) {
          headerMatrix.push(row);
          for (let k = 1; k < axisCellLength; k++) {
            let extraRow = Array(matrix[0].length - 1).fill({
              text: "",
              style: {},
            });
            extraRow.push({ text: domain[k], style: {} });
            headerMatrix.push(extraRow);
          }
          continue rowMatrixIter;
        }
      }
    }
    headerMatrix.push(row);
  }
  return headerMatrix;
};

const extractPivotData = (canvas) => {
  let columnWidth = [0];
  let columnAxisLength = [];

  // Exporting data from the column matrix
  const columnMatrix = canvas._composition.layout._columnMatrix._layoutMatrix;
  const topMatrix = canvas._composition.layout._columnMatrix._primaryMatrix;
  const bottomMatrix =
    canvas._composition.layout._columnMatrix._secondaryMatrix;

  let columnHeaders;
  let topHeaders = new Array(topMatrix.length);
  let bottomHeaders = new Array(bottomMatrix.length);

  topHeaders = generateTBMatrix(topMatrix, topHeaders, {
    columnWidth: columnWidth,
    columnAxisLength: columnAxisLength,
  });

  bottomHeaders = generateTBMatrix(bottomMatrix, bottomHeaders, {
    columnWidth: columnWidth,
    columnAxisLength: columnAxisLength,
  });

  columnHeaders = [topHeaders, bottomHeaders];
  console.log("columnHeaders", columnHeaders);

  // Exporting data from the row matrix
  const rowMatrix = canvas._composition.layout._rowMatrix._layoutMatrix;
  const leftMatrix = canvas._composition.layout._rowMatrix._primaryMatrix;
  const rightMatrix = canvas._composition.layout._rowMatrix._secondaryMatrix;

  let rowHeaders;
  let leftHeaders = [];
  let rightHeaders = [];

  let rowWidth = [0];
  let rowAxisLength = [];

  leftHeaders = generateLRMatrix(leftMatrix, leftHeaders, {
    rowWidth: rowWidth,
    rowAxisLength: rowAxisLength,
  });

  rightHeaders = generateLRMatrix(rightMatrix, rightHeaders, {
    rowWidth: rowWidth,
    rowAxisLength: rowAxisLength,
  });

  rowHeaders = [leftHeaders, rightHeaders];

  console.log("rowHeaders", rowHeaders);
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

  let geomData = Array.from({ length: rowWidth[0] }, () =>
    Array(columnWidth[0]).fill(null)
  );

  let prevY = 0;
  for (let i = 0; i < geomMatrix[0].length; i++) {
    let prevX = 0;
    for (let j = 0; j < geomMatrix.length; j++) {
      let textLayer = geomMatrix[j][i]._source._layers[0];
      let dataLength;
      let axes = [textLayer._axes.x, textLayer._axes.y];
      let data = geomMatrix[j][i]._source._layers[0]._pointMap;
      dataLength = Object.keys(data).length;

      if (Object.keys(data).length === 0) {
        data = geomMatrix[j][i]._source._layers[0]._normalizedData[0];
        dataLength = data.length;
      }

      for (let k = 0; k < dataLength; k++) {
        const dataPoint = data[k];

        if (!("xIndex" in dataPoint)) {
          dataPoint.xIndex = axes[0].getIndex(dataPoint.x);
        }
        if (!("yIndex" in dataPoint)) {
          dataPoint.yIndex = axes[1].getIndex(dataPoint.y);
        }

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
