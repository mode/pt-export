import {
  STARTCOLUMN,
  // STARTROW,
  PADDING,
  STARTROW,
  TABLEBORDERCOLOR,
  TABLEBORDERSTYLE,
  //   CELLBORDERCOLOR,
  //   CELLBORDERCOLOR
} from "../../constants/layoutConstants";

export const setColumnWidth = (maxWidths, sheet) => {
  maxWidths.forEach((width, index) => {
    sheet.getColumn(index + STARTCOLUMN).width = width + PADDING; // ExcelJS columns are 1-based
  });
};

export const applyZebraStriping = (targetCell) => {
  targetCell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFEEEEEE" }, // Light grey
  };
};

export const setAlignment = (targetCell, verticalAlign, horizontalAlign) => {
  targetCell.alignment = {
    wrapText: true,
    horizontal: horizontalAlign ? horizontalAlign : "left",
    vertical: verticalAlign,
  };
};

export const setBorder = (targetCell, borderPosition, style, color) => {
  switch (borderPosition) {
    case "left":
      targetCell.border = {
        ...targetCell.border,
        left: { style: style, color: { argb: color } },
      };
      break;

    case "right":
      targetCell.border = {
        ...targetCell.border,
        right: { style: style, color: { argb: color } },
      };
      break;

    case "top":
      targetCell.border = {
        ...targetCell.border,
        top: { style: style, color: { argb: color } },
      };
      break;

    case "bottom":
      targetCell.border = {
        ...targetCell.border,
        bottom: { style: style, color: { argb: color } },
      };
    default:
  }
};

export const applyTableBorder = (
  sheet,
  columnHeaderContent,
  totalColumns,
  totalRows
) => {
  let startRow = STARTROW;
  if (columnHeaderContent) {
    totalRows += 1;
    startRow -= 1;
  }

  //Applying borders to the top
  for (let colIdx = 0; colIdx < totalColumns; colIdx++) {
    const cellRef = sheet.getCell(startRow, STARTCOLUMN + colIdx);
    setBorder(cellRef, "top", TABLEBORDERSTYLE, TABLEBORDERCOLOR);
  }

  //Applying borders to the left
  for (let rowIdx = 0; rowIdx < totalRows; rowIdx++) {
    const cellRef = sheet.getCell(startRow + rowIdx, STARTCOLUMN);
    setBorder(cellRef, "left", TABLEBORDERSTYLE, TABLEBORDERCOLOR);
  }

  //Applying borders to the right
  for (let rowIdx = 0; rowIdx < totalRows; rowIdx++) {
    const cellRef = sheet.getCell(
      startRow + rowIdx,
      STARTCOLUMN + totalColumns - 1
    );
    setBorder(cellRef, "right", TABLEBORDERSTYLE, TABLEBORDERCOLOR);
  }

  //Applying borders to the bottom
  for (let colIdx = 0; colIdx < totalColumns; colIdx++) {
    const cellRef = sheet.getCell(
      startRow + totalRows - 1,
      STARTCOLUMN + colIdx
    );
    setBorder(cellRef, "bottom", TABLEBORDERSTYLE, TABLEBORDERCOLOR);
  }
};

export const applyColor = (cellRef, colors) => {
  cellRef.font = {
    color: { argb: colors[0] },
  };
};
