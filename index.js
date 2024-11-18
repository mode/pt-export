const convertToExcel = (data) => {
  console.log("Converting to excel");
};

const extractPivotData = (canvas) => {
  // Exporting data from the column matrix
  const columnMatrix = canvas._composition.layout._columnMatrix._tree.matrix;
  let columnHeaders = new Array(columnMatrix.length);

  columnMatrix.forEach((row, i) => {
    let prev = 0;
    let prevHeader = "";
    columnHeaders[i] = row.flatMap((cell, j) => {
      if (cell._source === null) {
        return "";
      } else if (typeof cell.source() === "object") {
        const axisCellLength = cell._source._domain.length;
        if (i > 0 && typeof columnMatrix[i - 1][j]._source === "string") {
          let k = i;

          while (k > 0) {
            columnHeaders[k - 1].splice(
              j + 1 + prev,
              0,
              ...Array(axisCellLength - 1).fill("")
            );
            k -= 1;
          }
          prev += axisCellLength - 1;
        }
        return cell._source._domain;
      } else {
        if (prevHeader === cell._source) {
          return "";
        } else {
          prevHeader = cell._source;
          return prevHeader;
        }
      }
    });
  });
  console.log("columnHeaders", columnHeaders);

  // Exporting data from the row matrix
  const rowMatrix = canvas._composition.layout._rowMatrix._tree.matrix;
  const rowMatLength = rowMatrix.length;
  const columnLength = rowMatrix[0].length;

  let rowHeaders = [];

  rowMatrixIter: for (let i = columnMatrix.length; i < rowMatLength; i++) {
    let row = [];
    for (let j = 0; j < columnLength; j++) {
      const cell = rowMatrix[i][j];
      if (cell._source === null) {
        row.push("");
      } else if (typeof cell._source === "string") {
        row.push(cell._source);
      } else {
        const domain = cell._source._domain.reverse();
        const axisCellLength = domain.length;

        row.push(domain[0]);
        if (axisCellLength > 1) {
          rowHeaders.push(row);
          for (let k = 1; k < axisCellLength; k++) {
            let extraRow = Array(columnLength - 1).fill("");
            extraRow.push(domain[k]);
            rowHeaders.push(extraRow);
            continue rowMatrixIter;
          }
        }
      }
    }
    rowHeaders.push(row);
  }
  console.log("rowHeaders", rowHeaders);
};

export const exportToExcel = (canvas) => {
  const data = extractPivotData(canvas);
  convertToExcel(data);
};
