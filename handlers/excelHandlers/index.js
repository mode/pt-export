import ExcelJS from "exceljs";
import {
  insertColumnHeaders,
  insertRowHeaders,
  insertGeomMatrix,
} from "./dataPopulator";
import { setColumnWidth } from "./helpers";


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

export const convertToExcel = (params) => {
  const { columnHeaders, rowHeaders, xSplit, ySplit, geomData, maxWidths } =
    params;

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
  insertColumnHeaders(columnHeaders[0], sheet, 0);

  //Insert left row matrix
  insertRowHeaders(rowHeaders[0], sheet, columnHeaders[0].length, 0, true);

  //Insert bottom column matrix, if present
  if (columnHeaders[1].length) {
    const prevRowsCount = columnHeaders[0].length + rowHeaders[0].length;
    insertColumnHeaders(columnHeaders[1], sheet, prevRowsCount);
  }

  insertGeomMatrix(
    geomData,
    sheet,
    columnHeaders[0].length,
    rowHeaders[0].length > 0 ? rowHeaders[0][0].length : 0
  );

  //Insert right row matrix, if present
  if (rowHeaders[1][0].length) {
    const prevColsCount = rowHeaders[0][1].length + geomData[0].length;
    insertRowHeaders(
      rowHeaders[1],
      sheet,
      columnHeaders[0].length,
      prevColsCount,
      false
    );
  }

  // Create a BLOB object from the workbook
  workbook.xlsx.writeBuffer().then((buffer) => {
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    saveExcelFile(blob);
  });
};
