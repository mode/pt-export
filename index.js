const convertToExcel = (data) => {
  console.log("Converting to excel");
};

const extractPivotData = (canvas) => {
  let columnWidth = 0;
  let columnAxisLength = [];

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
        columnWidth += axisCellLength;
        columnAxisLength.push(axisCellLength);

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
  let rowWidth = 0;
  let rowAxisLength = [];

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
        rowWidth += axisCellLength;
        rowAxisLength.push(axisCellLength);

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
  console.log("columnWidth", columnWidth);

  console.log("rowWidth", rowWidth);
  console.log("columnAxisLength", columnAxisLength);
  console.log("rowAxisLength", rowAxisLength);

  // Exporting data from the geom matrix
  const geomMatrix = canvas._composition.layout._centerMatrix._layoutMatrix;

  let geomData = Array.from({ length: rowWidth }, () =>
    Array(columnWidth).fill(null)
  );

  let prevY = 0;
  for (let i = 0; i < geomMatrix[0].length; i++) {
    let prevX = 0;
    for (let j = 0; j < geomMatrix.length; j++) {
      let data = geomMatrix[j][i]._source._layers[0]._pointMap;

      for (let k = 0; k < Object.keys(data).length; k++) {
        const dataPoint = data[k];

        if (
          dataPoint.yIndex + prevX < rowWidth &&
          dataPoint.xIndex + prevY < columnWidth
        ) {
          if (
            geomData[dataPoint.yIndex + prevX][dataPoint.xIndex + prevY] ===
            null
          ) {
            geomData[dataPoint.yIndex + prevX][dataPoint.xIndex + prevY] = [];
          }
          geomData[dataPoint.yIndex + prevX][dataPoint.xIndex + prevY].push(
            dataPoint
          );
        }
      }
      prevX += rowAxisLength[j];
    }
    prevY += columnAxisLength[i];
  }
  console.log("geomData", geomData);
};

export const exportToExcel = (canvas) => {
  const data = extractPivotData(canvas);
  convertToExcel(data);
};
