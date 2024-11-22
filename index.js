const convertToExcel = (data) => {
  console.log("Converting to excel");
};

const generateTBMatrix = (matrix, headerMatrix, columnData) => {
  let { columnWidth, columnAxisLength } = columnData;

  matrix.forEach((row, i) => {
    let prev = 0;
    let prevHeader = "";
    headerMatrix[i] = row.flatMap((cell, j) => {
      if (cell._source === null) {
        return { text: "", style: {} };
      } else if (typeof cell.source() === "object") {
        const axisCellLength = cell._source._domain.length;
        columnWidth[0] += axisCellLength;
        columnAxisLength.push(axisCellLength);

        if (i > 0 && typeof matrix[i - 1][j]._source === "string") {
          let k = i;

          while (k > 0) {
            headerMatrix[k - 1].splice(
              j + 1 + prev,
              0,
              ...Array(axisCellLength - 1).fill({ text: "", style: {} })
            );
            k -= 1;
          }
          prev += axisCellLength - 1;
        }
        return cell._source._domain.map((element) => ({
          text: element,
          style: {},
        }));
      } else {
        if (prevHeader === cell._source) {
          return { text: "", style: {} };
        } else {
          prevHeader = cell._source;
          return { text: prevHeader, style: {} };
        }
      }
    });
  });

  return headerMatrix;
};

const generateLRMatrix = (matrix, headerMatrix, rowData) => {
  let { rowWidth, rowAxisLength, extraCellLengths } = rowData;

  rowMatrixIter: for (
    let i = extraCellLengths[0];
    i < matrix.length - extraCellLengths[1];
    i++
  ) {
    let row = [];
    for (let j = 0; j < matrix[0].length; j++) {
      const cell = matrix[i][j];
      if (cell._source === null) {
        row.push({ text: "", style: {} });
      } else if (typeof cell._source === "string") {
        row.push({ text: cell._source, style: {} });
      } else {
        const domain = cell._source._domain.reverse();
        const axisCellLength = domain.length;
        rowWidth[0] += axisCellLength;
        rowAxisLength.push(axisCellLength);

        row.push({ text: domain[0], style: {} });
        if (axisCellLength > 1) {
          headerMatrix.push(row);
          for (let k = 1; k < axisCellLength; k++) {
            let extraRow = Array(matrix[0].length - 1).fill({
              text: "",
              style: {},
            });
            extraRow.push({ text: domain[k], style: {} });
            headerMatrix.push(extraRow);
          }
          continue rowMatrixIter;
        }
      }
    }
    headerMatrix.push(row);
  }
  return headerMatrix;
};

const extractPivotData = (canvas) => {
  let columnWidth = [0];
  let columnAxisLength = [];

  // Exporting data from the column matrix
  const columnMatrix = canvas._composition.layout._columnMatrix._layoutMatrix;
  const topMatrix = canvas._composition.layout._columnMatrix._primaryMatrix;
  const bottomMatrix =
    canvas._composition.layout._columnMatrix._secondaryMatrix;

  let columnHeaders;
  let topHeaders = new Array(topMatrix.length);
  let bottomHeaders = new Array(bottomMatrix.length);

  topHeaders = generateTBMatrix(topMatrix, topHeaders, {
    columnWidth: columnWidth,
    columnAxisLength: columnAxisLength,
  });

  bottomHeaders = generateTBMatrix(bottomMatrix, bottomHeaders, {
    columnWidth: columnWidth,
    columnAxisLength: columnAxisLength,
  });

  columnHeaders = [topHeaders, bottomHeaders];
  console.log("columnHeaders", columnHeaders);
  console.log(canvas._composition.layout);
  // Exporting data from the row matrix
  const matrix = canvas._composition.layout._rowMatrix;
  const rowMatrix = matrix._layoutMatrix;
  const leftMatrix = matrix._primaryMatrix;
  const rightMatrix = matrix._secondaryMatrix;
  const extraCellLengths = matrix._config.extraCellLengths;

  let rowHeaders;
  let leftHeaders = [];
  let rightHeaders = [];

  let rowWidth = [0];
  let rowAxisLength = [];

  leftHeaders = generateLRMatrix(leftMatrix, leftHeaders, {
    rowWidth: rowWidth,
    rowAxisLength: rowAxisLength,
    extraCellLengths: extraCellLengths,
  });

  rightHeaders = generateLRMatrix(rightMatrix, rightHeaders, {
    rowWidth: rowWidth,
    rowAxisLength: rowAxisLength,
    extraCellLengths: extraCellLengths,
  });

  rowHeaders = [leftHeaders, rightHeaders];

  console.log("rowHeaders", rowHeaders);

  // Exporting data from the geom matrix
  const geomMatrix = canvas._composition.layout._centerMatrix._layoutMatrix;

  let geomData = Array.from({ length: rowWidth[0] }, () =>
    Array(columnWidth[0]).fill(null)
  );

  let prevY = 0;
  for (let i = 0; i < geomMatrix[0].length; i++) {
    let prevX = 0;
    for (let j = 0; j < geomMatrix.length; j++) {
      let textLayer = geomMatrix[j][i]._source._layers[0];
      let dataLength;
      let axes = [textLayer._axes.x, textLayer._axes.y];
      let data = geomMatrix[j][i]._source._layers[0]._pointMap;
      dataLength = Object.keys(data).length;

      if (Object.keys(data).length === 0) {
        data = geomMatrix[j][i]._source._layers[0]._normalizedData[0];
        dataLength = data.length;
      }

      for (let k = 0; k < dataLength; k++) {
        const dataPoint = data[k];

        if (!("xIndex" in dataPoint)) {
          dataPoint.xIndex = axes[0].getIndex(dataPoint.x);
        }
        if (!("yIndex" in dataPoint)) {
          dataPoint.yIndex = axes[1].getIndex(dataPoint.y);
        }

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
          //   console.log(canvas._composition.layout._centerMatrix);
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

  const xSplit = rowMatrix[0].length;
  const ySplit = columnMatrix.length;
  console.log("xSplit", xSplit);
  console.log("ySplit", ySplit);
};

export const exportToExcel = (canvas) => {
  const data = extractPivotData(canvas);
  convertToExcel(data);
};
