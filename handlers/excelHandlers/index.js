import ExcelJS from "exceljs";
import {
  insertColumnMatrix,
  insertRowMatrix,
  insertGeomMatrix,
  insertColumnHeader,
} from "./dataPopulator";
import {setColumnWidth, applyTableBorder } from "./helpers";
import {
  STARTCOLUMN,
  STARTROW,
  TABLEBORDERCOLOR,
  TABLEBORDERSTYLE,
} from "../../constants/layoutConstants";

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

export const downloadSheet = (workBook) => {
    // Create a BLOB object from the workbook
    workBook.xlsx.writeBuffer().then((buffer) => {
      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      saveExcelFile(blob);
    });

}
export const convertToExcel = (params) => {
  const {
    columnHeaders: columnMatrix,
    rowHeaders: rowMatrix,
    xSplit,
    ySplit,
    geomData,
    maxWidths,
    columnHeaderContent,
    hasBg,
    hasColor,
  } = params;

  const workbook = new ExcelJS.Workbook();

  const sheet = workbook.addWorksheet("Sheet", {
    // views: [
    //   {
    //     state: "frozen", // Enables frozen view
    //     xSplit: STARTCOLUMN + xSplit - 1, // Freezes column matrix
    //     ySplit: STARTROW + ySplit - 1, // Freezes rows
    //   },
    // ],
  });

  setColumnWidth(maxWidths, sheet);

  //Insert top column Matrix
  insertColumnMatrix(columnMatrix[0], sheet, 0);

  //Insert left row matrix
  insertRowMatrix(rowMatrix[0], sheet, columnMatrix[0].length, 0, true);

  //Insert bottom column matrix, if present
  if (columnMatrix[1].length) {
    const prevRowsCount = columnMatrix[0].length + rowMatrix[0].length;
    insertColumnMatrix(columnMatrix[1], sheet, prevRowsCount);
  }

  insertGeomMatrix(
    geomData,
    sheet,
    columnMatrix[0].length,
    rowMatrix[0].length > 0 ? rowMatrix[0][0].length : 0,
    hasBg,
    hasColor
  );

  const totalColumns =
    rowMatrix[0][0].length + geomData[0].length + rowMatrix[1][0].length;

  const totalRows = columnMatrix[0].length + columnMatrix[1].length + geomData.length

  if (columnHeaderContent)
    insertColumnHeader(sheet, totalColumns, columnHeaderContent);

  //Insert right row matrix, if present
  if (rowMatrix[1][0].length) {
    const prevColsCount = rowMatrix[0][1].length + geomData[0].length;
    insertRowMatrix(
      rowMatrix[1],
      sheet,
      columnMatrix[0].length,
      prevColsCount,
      false
    );
  }

  applyTableBorder(sheet, columnHeaderContent, totalColumns, totalRows);

  return workbook;
};
