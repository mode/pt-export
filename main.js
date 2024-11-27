import { extractPivotData } from "./handlers/muzeHandlers/index.js";
import { convertToExcel } from "./handlers/excelHandlers/index.js";

export const exportToExcel = (canvas) => {
  const data = extractPivotData(canvas);
  convertToExcel(data);
};
