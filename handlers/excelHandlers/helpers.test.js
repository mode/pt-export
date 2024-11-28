import { expect, test } from 'vitest'
import {setAlignment} from "./helpers.js";
import ExcelJS from "exceljs";

test('testing', () => {
    // const target
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

    const targetCell = sheet.getCell('E3');
    setAlignment(targetCell, "top")
    expect(targetCell.alignment).toStrictEqual({
        wrapText: true,
        horizontal: "left",
        vertical: "top",
      })
  })
