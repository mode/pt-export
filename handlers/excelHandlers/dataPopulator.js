import { isValidHslaFormat } from "../../colorUtils/colorValidator";
import { hslaToArgb } from "../../colorUtils/colorFormatter";
import {
  STARTCOLUMN,
  STARTROW,
  //   PADDING,
  TABLEBORDER,
} from "../../constants/layoutConstants";
import { applyZebraStriping, setAlignment, setBorder } from "./helpers";

export const insertColumnHeaders = (columnHeaders, sheet, prevRowsCount) => {
  //Adding columnHeaders data to the sheet
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

  //Merging cells, when necessary
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
      );

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

export const insertRowHeaders = (
  rowHeaders,
  sheet,
  prevRowsCount,
  prevColsCount,
  isLeftMatrix
) => {
  const rowMatrixLen = rowHeaders[0].length;

  //Adding rowHeaders data to the sheet and applying zebra striping only to the axis of left row matrix, not facets
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
        STARTCOLUMN + colIdx + prevColsCount
      );
      targetCell.value = formattedCell;
      if (isLeftMatrix) {
        if (colIdx + 1 === rowMatrixLen && rowIdx % 2 === 1) {
          applyZebraStriping(targetCell);
          setBorder(targetCell, "left", "thin", "D9D9D9");
        }
        //Applying border to the table
        if (colIdx === 0) {
          setBorder(targetCell, "left", "medium", TABLEBORDER);
        }
      }
    });
  });

  const numOfCols = rowHeaders[0].length;
  const rowLen = rowHeaders.length;

  //Merging cells, when necessary
  for (let col = 1; col < numOfCols; col++) {
    //No merging for the axis labels
    let startRow = null;
    let mergeValue = "";

    for (let row = 1; row <= rowLen; row++) {
      const cellValue = rowHeaders[row - 1][col - 1].text;

      if (cellValue && cellValue !== mergeValue) {
        // If a new mergeValue is found, close the previous merge (if any)
        if (startRow !== null) {
          sheet.mergeCells(
            prevRowsCount + startRow + STARTROW - 1,
            STARTCOLUMN + col + prevColsCount - 1,
            prevRowsCount + STARTROW + row - 2,
            STARTCOLUMN + col + prevColsCount - 1
          );

          const targetCell = sheet.getCell(
            prevRowsCount + startRow + STARTROW - 1,
            STARTCOLUMN + col + prevColsCount - 1
          );
          setAlignment(targetCell, "top");
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
        STARTCOLUMN + col + prevColsCount - 1,
        prevRowsCount + rowLen + STARTROW - 1,
        STARTCOLUMN + col + prevColsCount - 1
      );

      const targetCell = sheet.getCell(
        prevRowsCount + startRow + STARTROW - 1,
        col + STARTCOLUMN + prevColsCount - 1
      );
      setAlignment(targetCell, "top");
    }
  }

  for (let i = 0; i < rowLen; i++) {
    const targetCell = sheet.getCell(
      prevRowsCount + STARTROW + i,
      STARTCOLUMN + rowHeaders[0].length - 1
    );
    setAlignment(targetCell, "middle");
  }
};

export const insertGeomMatrix = (
  geomData,
  sheet,
  prevRowsCount,
  prevColCount
) => {
  let rowNo = prevRowsCount + STARTROW;
  const columnNoStart = prevColCount + STARTCOLUMN;

  geomData.forEach((row, rowIdx) => {
    let columnNo = columnNoStart;

    row.forEach((cell, idx) => {
      const cellRef = sheet.getCell(rowNo, columnNo);
      let align = "left";

      if (cell === null) {
        // Skip the cell if it's null
      } else {
        const colors = [];
        const values = [];
        cell.forEach((item) => {
          let value = item.text;
          let color = item.color;

          let argbColor = "000000"; //Default to black
          let isNumeric = false;

          if (isValidHslaFormat(color)) {
            argbColor = hslaToArgb(color);
          }

          if (Number(value) || Number(value) === 0) {
            align = "right";
            isNumeric = true;
          }

          colors.push(argbColor);
          values.push(isNumeric ? Number(value) : value);
        });

        if (values.length > 1) {
          const richTextArray = values.map((value, index) => ({
            text: value.toString() + " ", // Adding a space after each number
            font: {
              color: { argb: colors[index] },
            },
          }));

          cellRef.value = { richText: richTextArray };
        } else {
          cellRef.value = values[0];

          cellRef.font = {
            color: { argb: colors[0] },
          };
        }

        setAlignment(cellRef, "top", align);
      }

      //Apply Zebra striping
      if (rowIdx % 2 === 1) {
        applyZebraStriping(cellRef);
        if (idx === geomData[0].length - 1)
          setBorder(cellRef, "right", "thin", "D9D9D9");
      }

      //Apply table border
      if (idx === geomData[0].length - 1) {
        setBorder(cellRef, "right", "medium", TABLEBORDER);
      }
      columnNo++;
    });
    rowNo++;
  });
};
