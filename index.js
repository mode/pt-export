const convertToExcel = (data) => {
  console.log("Converting to excel");
};

const generateTBMatrix = (
  matrix,
  topMatrix,
  columnData,
  columnHeaders,
  columnMaxWidths
) => {
  let { columnWidth, columnAxisLength } = columnData;

  matrix.forEach((row, i) => {
    let prev = 0;
    let prevHeader = "";

    columnHeaders[i] = row.flatMap((cell, j) => {
      if (cell._source === null) {
        return { text: "", style: {} };
      } else if (typeof cell.source() === "object") {
        const axisCellLength = cell._source._domain.length;
        columnWidth[0] += axisCellLength;
        columnAxisLength.push(axisCellLength);

        if (i > 0 && typeof matrix[i - 1][j]._source === "string") {
          let k = i;

          while (k > 0) {
            columnHeaders[k - 1].splice(
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

  let topHeaders = columnHeaders.slice(0, topMatrix.length);
  let bottomHeaders = columnHeaders.slice(
    topMatrix.length,
    topMatrix.length + columnHeaders.length
  );

  return [topHeaders, bottomHeaders];
};

const generateLRMatrix = (
  matrix,
  leftMatrix,
  rightMatrix,
  rowData,
  rowHeaders,
  leftRowMaxWidths,
  rightRowMaxWidths
) => {
  let { rowWidth, rowAxisLength, extraCellLengths } = rowData;
  let rowMaxWidths = new Array(matrix[0].length).fill(0);

  if (rightMatrix.length > 0) {
    matrix = matrix.map((row) => row.reverse());
  }

  rowMatrixIter: for (
    let i = extraCellLengths[0];
    i < matrix.length - extraCellLengths[1];
    i++
  ) {
    let row = [];
    for (let j = 0; j < matrix[0].length; j++) {
      let cell = matrix[i][j];

      if (cell._source === null) {
        row.push({ text: "", style: {} });
      } else if (typeof cell._source === "string") {
        rowMaxWidths[j] = Math.max(rowMaxWidths[j], cell._source.length);
        row.push({ text: cell._source, style: {} });
      } else {
        const domain = cell._source._domain.reverse();
        const axisCellLength = domain.length;
        rowWidth[0] += axisCellLength;
        rowAxisLength.push(axisCellLength);

        rowMaxWidths[j] = Math.max(rowMaxWidths[j], domain[0].length);
        row.push({ text: domain[0], style: {} });

        if (axisCellLength > 1) {
          rowHeaders.push(row);

          for (let k = 1; k < axisCellLength; k++) {
            let extraRow = Array(matrix[0].length - 1).fill({
              text: "",
              style: {},
            });

            rowMaxWidths[j] = Math.max(rowMaxWidths[j], domain[k].length);
            extraRow.push({ text: domain[k], style: {} });
            rowHeaders.push(extraRow);
          }
          continue rowMatrixIter;
        }
      }
    }
    rowHeaders.push(row);
  }

  const leftHeaders = rowHeaders.map((row) =>
    row.slice(0, leftMatrix[0].length)
  );

  leftRowMaxWidths.push(...rowMaxWidths.slice(0, leftMatrix[0].length));

  const rightHeaders = rowHeaders.map((row) =>
    row.slice(leftMatrix[0].length, leftMatrix[0].length + matrix[0].length)
  );

  rightRowMaxWidths.push(
    ...rowMaxWidths.slice(
      leftMatrix[0].length,
      leftMatrix[0].length + matrix[0].length
    )
  );

  if (rightMatrix.length > 0) {
    let temp = rightRowMaxWidths;

    rightRowMaxWidths = leftRowMaxWidths;
    leftRowMaxWidths = temp;

    return [rightHeaders, leftHeaders];
  }

  return [leftHeaders, rightHeaders];
};

const extractPivotData = (canvas) => {
  let columnWidth = [0];
  let columnAxisLength = [];

  // Exporting data from the column matrix
  const columnMatrix = canvas._composition.layout._columnMatrix._layoutMatrix;
  const topMatrix = canvas._composition.layout._columnMatrix._primaryMatrix;
  const bottomMatrix =
    canvas._composition.layout._columnMatrix._secondaryMatrix;

  let cHeaders = new Array(columnMatrix.length);
  let columnMaxWidths = new Array(
    columnMatrix[columnMatrix.length - 1].length
  ).fill(0);

  let columnHeaders = generateTBMatrix(
    columnMatrix,
    topMatrix,
    {
      columnWidth: columnWidth,
      columnAxisLength: columnAxisLength,
    },
    cHeaders,
    columnMaxWidths
  );

  console.log("columnHeaders", columnHeaders);
  console.log("columnMaxWidths", columnMaxWidths);
  console.log(canvas._composition.layout);

  // Exporting data from the row matrix
  const matrix = canvas._composition.layout._rowMatrix;
  const rowMatrix = matrix._layoutMatrix;
  const leftMatrix = matrix._primaryMatrix;
  const rightMatrix = matrix._secondaryMatrix;
  const extraCellLengths = matrix._config.extraCellLengths;

  let rHeaders = [];

  let rowWidth = [0];
  let rowAxisLength = [];
  let leftRowMaxWidths = [];
  let rightRowMaxWidths = [];

  let rowHeaders = generateLRMatrix(
    rowMatrix,
    leftMatrix,
    rightMatrix,
    {
      rowWidth: rowWidth,
      rowAxisLength: rowAxisLength,
      extraCellLengths: extraCellLengths,
    },
    rHeaders,
    leftRowMaxWidths,
    rightRowMaxWidths
  );

  console.log("rowHeaders", rowHeaders);

  // Calculating the maximum width of text in each column
  const lastColumn =
    columnHeaders[1].length !== 0
      ? columnHeaders[1][columnHeaders[1].length - 1]
      : columnHeaders[0][columnHeaders[0].length - 1];

  const maxWidths = new Array(lastColumn.length).fill(0);
  let rightRowLength = rightRowMaxWidths.length;

  let lastInd = 0;
  lastColumn.forEach((obj, i) => {
    if (i < leftRowMaxWidths.length) {
      maxWidths[i] = Math.max(leftRowMaxWidths[i], obj.text.length);
    } else if (i >= lastColumn.length - rightRowLength) {
      maxWidths[i] = Math.max(rightRowMaxWidths[lastInd], obj.text.length);
      lastInd += 1;
    } else {
      maxWidths[i] = obj.text.length;
    }
  });

  console.log("maxWidths", maxWidths);

  // Exporting data from the geom matrix
  const geomMatrix = canvas._composition.layout._centerMatrix._layoutMatrix;

  let geomData = Array.from({ length: rowWidth[0] }, () =>
    Array(columnWidth[0]).fill(null)
  );

  const retinalAxes = geomMatrix[0][0]._source._axes;

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
        dataLength = data?.length;
      }

      for (let k = 0; k < dataLength; k++) {
        const dataPoint = data[k];

        if (dataPoint && !("xIndex" in dataPoint)) {
          dataPoint.xIndex = axes[0].getIndex(dataPoint.x);
        }
        if (dataPoint && !("yIndex" in dataPoint)) {
          dataPoint.yIndex = axes[1].getIndex(dataPoint.y);
        }

        if (
          dataPoint?.yIndex + prevX < rowWidth &&
          dataPoint?.xIndex + prevY < columnWidth
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

  const xSplit = rowHeaders[0][0].length;
  const ySplit = columnHeaders[0].length;
  console.log("xSplit", xSplit);
  console.log("ySplit", ySplit);
};

export const exportToExcel = (canvas) => {
  const data = extractPivotData(canvas);
  convertToExcel(data);
};
