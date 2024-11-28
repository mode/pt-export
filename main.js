import { extractPivotData } from "./handlers/muzeHandlers/index.js";
import { convertToExcel , downloadSheet } from "./handlers/excelHandlers/index.js";

export const prepareWorkSheet = (canvas) => {
  const data = extractPivotData(canvas);
  const workBook = convertToExcel(data);
  return workBook;
};

export const exportToExcel = (canvas) => {
   downloadSheet(prepareWorkSheet(canvas));
}
