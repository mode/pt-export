import {
  STARTCOLUMN,
  // STARTROW,
  PADDING,
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
        left: { style: style, color: { argb: color } },
      };
      break;

    case "right":
      targetCell.border = {
        right: { style: style, color: { argb: color } },
      };
      break;
    default:
  }
};
