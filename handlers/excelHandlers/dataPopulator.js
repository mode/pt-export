import { isValidHslaFormat } from "../../colorUtils/colorValidator";
import { hslaToArgb } from "../../colorUtils/colorFormatter";
import {
  CELLBORDERCOLOR,
  CELLBORDERSTYLE,
  STARTCOLUMN,
  STARTROW,
  //   PADDING,
} from "../../constants/layoutConstants";
import {
  applyZebraStriping,
  setAlignment,
  setBorder,
  applyColor,
} from "./helpers";

export const insertColumnMatrix = (columnMatrix, sheet, prevRowsCount) => {
  //Adding columnMatrix data to the sheet
  columnMatrix.forEach((row, rowIdx) => {
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
  columnMatrix.forEach((rowData, rowIdx) => {
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

export const insertRowMatrix = (
  rowMatrix,
  sheet,
  prevRowsCount,
  prevColsCount,
  isLeftMatrix
) => {
  const rowMatrixLen = rowMatrix[0].length;

  //Adding rowMatrix data to the sheet and applying zebra striping only to the axis of left row matrix, not facets
  rowMatrix.forEach((row, rowIdx) => {
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
          setBorder(targetCell, "left", CELLBORDERSTYLE, CELLBORDERCOLOR);
          setBorder(targetCell, "right", CELLBORDERSTYLE, CELLBORDERCOLOR);
        }
      }
    });
  });

  const numOfCols = rowMatrix[0].length;
  const rowLen = rowMatrix.length;

  //Merging cells, when necessary
  for (let col = 1; col < numOfCols; col++) {
    //No merging for the axis labels
    let startRow = null;
    let mergeValue = "";

    for (let row = 1; row <= rowLen; row++) {
      const cellValue = rowMatrix[row - 1][col - 1].text;

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
      STARTCOLUMN + rowMatrix[0].length - 1
    );
    setAlignment(targetCell, "middle");
  }
};

export const addValue = (item, alignment) => {
  let value = item.text;
  let isNumeric = false;

  if (Number(value) || Number(value) === 0) {
    alignment.align = "right";
    isNumeric = true;
  }

  return isNumeric ? Number(value) : value;
};

export const addColor = (item, hasBg, hasColor) => {
  let color = item.color;

  //Currently accepting only validHslaFormat
  if (isValidHslaFormat(color)) {
    color = hslaToArgb(color);
  } else {
    return;
  }

  let backgroundColor = item.backgroundColor;

  //Currently accepting only validHslaFormat
  if (isValidHslaFormat(backgroundColor)) {
    backgroundColor = hslaToArgb(backgroundColor);
  } else {
    return;
  }

  if (!hasBg && !hasColor) {
    //both color and backgroundColor encoding is not present

    //default black color for the text
    return "000000";
  } else if (hasBg && !hasColor) {
    //only backgroundColor encoding is present, not the color encoding

    // default white color for the text
    return backgroundColor;
  } else {
    //Both color and backgroundColor encoding is present
    //only color encoding is present, not the backgroundColor encoding

    return color;
  }
};

export const insertGeomMatrix = (
  geomData,
  sheet,
  prevRowsCount,
  prevColCount,
  hasBg,
  hasColor
) => {
  debugger;
  let rowNo = prevRowsCount + STARTROW;
  const columnNoStart = prevColCount + STARTCOLUMN;

  geomData.forEach((row, rowIdx) => {
    let columnNo = columnNoStart;

    row.forEach((cell, idx) => {
      const cellRef = sheet.getCell(rowNo, columnNo);
      let alignment = { align: "left" };

      if (cell === null) {
        // Skip the cell if it's null
      } else {
        const colors = [];
        const values = [];

        cell.forEach((item) => {
          values.push(addValue(item, alignment));

          colors.push(addColor(item, hasBg, hasColor));
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

          applyColor(cellRef, colors);
        }

        setAlignment(cellRef, "top", alignment.align);
      }

      //Apply Zebra striping
      if (rowIdx % 2 === 1) {
        applyZebraStriping(cellRef);
        if (idx === geomData[0].length - 1)
          setBorder(cellRef, "right", CELLBORDERSTYLE, CELLBORDERCOLOR); //Light grey border on the right
      }

      columnNo++;
    });
    rowNo++;
  });
};

export const insertColumnHeader = (
  sheet,
  totalColumns,
  columnHeaderContent
) => {
  debugger;
  const middleCell = sheet.getCell(
    Math.max(STARTROW - 1, 1), // Ensure the row is at least 1
    STARTCOLUMN + Math.floor(totalColumns / 2) // Use Math.floor to get an integer column index
  );
  middleCell.value = columnHeaderContent;
};
