import ExcelJS from "exceljs";
import { STARTCOLUMN, STARTROW } from "./constants.js";

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

const insertColumnHeaders = (columnHeaders, sheet, prevRowsCount) => {
  columnHeaders.forEach((row, rowIdx) => {
    row.forEach((cell, colIdx) => {
      const formattedCell =
        cell && typeof cell === "object"
          ? Number(cell.text)
            ? Number(cell.text)
            : cell.text
          : " ";

      const targetCell = sheet.getCell(
        STARTROW + rowIdx + prevRowsCount,
        STARTCOLUMN + colIdx
      );
      targetCell.value = formattedCell;
    });
  });

  columnHeaders.forEach((rowData, rowIdx) => {
    let startIdx = 0;
    while (startIdx < rowData.length && rowData[startIdx].text === "")
      startIdx++;

    if (startIdx > 0)
      sheet.mergeCells(
        prevRowsCount + rowIdx + STARTROW,
        STARTCOLUMN,
        prevRowsCount + rowIdx + STARTROW,
        STARTCOLUMN + startIdx - 1
      ); //Need to pass the rowNo and columnNo
    if (startIdx === rowData.length) return;

    let endIdx = startIdx;
    while (startIdx < rowData.length) {
      if (rowData[startIdx].text !== "" && startIdx !== endIdx) {
        sheet.mergeCells(
          prevRowsCount + rowIdx + STARTROW,
          endIdx + STARTCOLUMN,
          prevRowsCount + rowIdx + STARTROW,
          startIdx + STARTCOLUMN - 1
        );
        endIdx = startIdx;
      }
      startIdx++;
    }
    sheet.mergeCells(
      prevRowsCount + rowIdx + STARTROW,
      endIdx + STARTCOLUMN,
      prevRowsCount + rowIdx + STARTROW,
      startIdx + STARTCOLUMN - 1 
    );
  });
};

const insertRowHeaders = (rowHeaders, sheet, prevRowsCount) => {
 rowHeaders.forEach((row, rowIdx) => {
    row.forEach((cell, colIdx) => {
      const formattedCell =
        cell && typeof cell === "object"
          ? Number(cell.text)
            ? Number(cell.text)
            : cell.text
          : " ";

      const targetCell = sheet.getCell(
        STARTROW + rowIdx + prevRowsCount,
        STARTCOLUMN + colIdx
      );
      targetCell.value = formattedCell;
    });
  });

  const numOfCols = rowHeaders[0].length;

  const rowLen = rowHeaders.length;
  // // Traverse columns
  for (let col = 1; col <= numOfCols; col++) {
    let startRow = null;
    let mergeValue = "";

    for (let row = 1; row <= rowLen; row++) {
      const cellValue = rowHeaders[row - 1][col - 1].text;

      if (cellValue && cellValue !== mergeValue) {
        // If a new mergeValue is found, close the previous merge (if any)
        if (startRow !== null) {
          sheet.mergeCells(
            prevRowsCount + startRow + STARTROW - 1,
            STARTCOLUMN + col - 1,
            prevRowsCount + STARTROW + row - 2,
            STARTCOLUMN + col - 1
          );

          sheet.getCell(prevRowsCount + startRow + STARTROW - 1, STARTCOLUMN + col - 1).alignment = {
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
        prevRowsCount + startRow + STARTROW - 1,
        STARTCOLUMN + col - 1,
        prevRowsCount + rowLen + STARTROW - 1,
        STARTCOLUMN + col - 1
      );

      sheet.getCell(prevRowsCount + startRow + STARTROW - 1, col + STARTCOLUMN - 1).alignment = {
        vertical: "top",
      };
    }
  }
};

const insertGeomMatrix = (geomData, sheet, prevRowsCount, prevColCount) => {
  let rowNo = prevRowsCount + STARTROW;
  const columnNoStart = prevColCount + STARTCOLUMN;

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

const setColumnWidth = (maxWidths, sheet) => {
  maxWidths.forEach((width, index) => {
    sheet.getColumn(index + STARTCOLUMN).width = width; // ExcelJS columns are 1-based
  });
};

const convertToExcel = (params) => {
  const { columnHeaders, rowHeaders, xSplit, ySplit, geomData, maxWidths } =
    params;
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
  debugger;
  setColumnWidth(maxWidths, sheet);

  insertColumnHeaders(columnHeaders[0], sheet, 0); //Insert top column Matrix

  insertRowHeaders(rowHeaders[0], sheet, columnHeaders[0].length); //Inserting left row matrix

  if (columnHeaders[1].length) {
    const prevRowsCount = columnHeaders[0].length + rowHeaders[0].length;
    insertColumnHeaders(columnHeaders[1], sheet, prevRowsCount); //Inserting bottom column matrix
  }

  insertGeomMatrix(
    geomData,
    sheet,
    columnHeaders[0].length,
    rowHeaders[0].length > 0 ? rowHeaders[0][0].length : 0
  );

  // Create a BLOB object from the workbook
  workbook.xlsx.writeBuffer().then((buffer) => {
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    saveExcelFile(blob);
  });
};

const generateTBMatrix = (
  matrix,
  topMatrix,
  columnData,
  columnHeaders,
  columnMaxWidths
) => {
  let { columnWidth, columnAxisLength } = columnData;

  matrix.forEach((row, i) => {
    let prev = 0;
    let prevHeader = "";

    columnHeaders[i] = row.flatMap((cell, j) => {
      if (cell._source === null) {
        return { text: "", style: {} };
      } else if (typeof cell.source() === "object") {
        const axisCellLength = cell._source._domain.length;
        columnWidth[0] += axisCellLength;
        columnAxisLength.push(axisCellLength);

        if (i > 0 && typeof matrix[i - 1][j]._source === "string") {
          let k = i;

          while (k > 0) {
            columnHeaders[k - 1].splice(
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

  let topHeaders = columnHeaders.slice(0, topMatrix.length);
  let bottomHeaders = columnHeaders.slice(
    topMatrix.length,
    topMatrix.length + columnHeaders.length
  );

  return [topHeaders, bottomHeaders];
};

const generateLRMatrix = (
  matrix,
  leftMatrix,
  rowData,
  rowHeaders,
  rowMaxWidths
) => {
  let { rowWidth, rowAxisLength, extraCellLengths } = rowData;
  debugger;

  rowMatrixIter: for (
    let i = extraCellLengths[0];
    i < matrix.length - extraCellLengths[1];
    i++
  ) {
    let row = [];
    for (let j = 0; j < matrix[0].length; j++) {
      const cell = matrix[i][j];
      if (cell._source === null) {
        row.push({ text: "", style: {} });
      } else if (typeof cell._source === "string") {
        rowMaxWidths[j] = Math.max(rowMaxWidths[j], cell._source.length);
        row.push({ text: cell._source, style: {} });
      } else {
        const domain = cell._source._domain.reverse();
        const axisCellLength = domain.length;
        rowWidth[0] += axisCellLength;
        rowAxisLength.push(axisCellLength);

        rowMaxWidths[j] = Math.max(rowMaxWidths[j], domain[0].length);
        row.push({ text: domain[0], style: {} });
        if (axisCellLength > 1) {
          rowHeaders.push(row);
          for (let k = 1; k < axisCellLength; k++) {
            let extraRow = Array(matrix[0].length - 1).fill({
              text: "",
              style: {},
            });

            rowMaxWidths[j] = Math.max(rowMaxWidths[j], domain[k].length);
            extraRow.push({ text: domain[k], style: {} });
            rowHeaders.push(extraRow);
          }
          continue rowMatrixIter;
        }
      }
    }
    rowHeaders.push(row);
  }

  const leftHeaders = rowHeaders.map((row) =>
    row.slice(0, leftMatrix[0].length)
  );
  const rightHeaders = rowHeaders.map((row) =>
    row.slice(leftMatrix[0].length, leftMatrix[0].length + matrix[0].length)
  );

  return [leftHeaders, rightHeaders];
};

const extractPivotData = (canvas) => {
  let columnWidth = [0];
  let columnAxisLength = [];
  // Exporting data from the column matrix
  const columnMatrix = canvas._composition.layout._columnMatrix._layoutMatrix;
  const topMatrix = canvas._composition.layout._columnMatrix._primaryMatrix;
  const bottomMatrix =
    canvas._composition.layout._columnMatrix._secondaryMatrix;

  let cHeaders = new Array(columnMatrix.length);
  let columnMaxWidths = new Array(
    columnMatrix[columnMatrix.length - 1].length
  ).fill(0);

  let columnHeaders = generateTBMatrix(
    columnMatrix,
    topMatrix,
    {
      columnWidth: columnWidth,
      columnAxisLength: columnAxisLength,
    },
    cHeaders,
    columnMaxWidths
  );

  console.log("columnHeaders", columnHeaders);
  console.log("columnMaxWidths", columnMaxWidths);
  console.log(canvas._composition.layout);

  // Exporting data from the row matrix
  const matrix = canvas._composition.layout._rowMatrix;
  const rowMatrix = matrix._layoutMatrix;
  const leftMatrix = matrix._primaryMatrix;
  const rightMatrix = matrix._secondaryMatrix;
  const extraCellLengths = matrix._config.extraCellLengths;

  let rHeaders = [];

  let rowWidth = [0];
  let rowAxisLength = [];
  let rowMaxWidths = new Array(rowMatrix[0].length).fill(0);

  let rowHeaders = generateLRMatrix(
    rowMatrix,
    leftMatrix,
    {
      rowWidth: rowWidth,
      rowAxisLength: rowAxisLength,
      extraCellLengths: extraCellLengths,
    },
    rHeaders,
    rowMaxWidths
  );

  console.log("rowHeaders", rowHeaders);
  //   console.log("rowMaxWidths", rowMaxWidths);

  // Generate max column widths
  const lastColumn =
    columnHeaders[1].length !== 0
      ? columnHeaders[1][columnHeaders[1].length - 1]
      : columnHeaders[0][columnHeaders[0].length - 1];

  const maxWidths = new Array(lastColumn.length).fill(0);

  lastColumn.forEach((obj, i) => {
    if (rowMaxWidths[i]) {
      maxWidths[i] = Math.max(rowMaxWidths[i], obj.text.length);
    } else {
      maxWidths[i] = obj.text.length;
    }
  });

  console.log("maxWidths", maxWidths);

  // Exporting data from the geom matrix
  const geomMatrix = canvas._composition.layout._centerMatrix._layoutMatrix;
  debugger;
  let geomData = Array.from({ length: rowWidth[0] }, () =>
    Array(columnWidth[0]).fill(null)
  );

  const retinalAxes = geomMatrix[0][0]._source._axes;

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

        if (dataPoint && !("xIndex" in dataPoint)) {
          dataPoint.xIndex = axes[0].getIndex(dataPoint.x);
        }
        if (dataPoint && !("yIndex" in dataPoint)) {
          dataPoint.yIndex = axes[1].getIndex(dataPoint.y);
        }

        if (
          dataPoint?.yIndex + prevX < rowWidth &&
          dataPoint?.xIndex + prevY < columnWidth
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

  const xSplit = rowHeaders[0][0].length;
  const ySplit = columnHeaders[0].length;
  return {
    columnHeaders: columnHeaders,
    rowHeaders: rowHeaders,
    xSplit: xSplit,
    ySplit: ySplit,
    geomData: geomData,
    maxWidths: maxWidths,
  };
};

export const exportToExcel = (canvas) => {
  const data = extractPivotData(canvas);
  convertToExcel(data);
};
