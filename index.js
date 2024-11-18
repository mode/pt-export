const convertToExcel = (data) => {
  console.log("Converting to excel");
};

const extractPivotData = (canvas) => {
  //   console.log("Hello canvas", canvas);

  const columnMatrix = canvas._composition.layout._columnMatrix._tree.matrix;
  //   console.log("column-matrix", columnMatrix);

  let columnHeaders = new Array(columnMatrix.length);
  columnMatrix.forEach((row, i) => {
    // console.log(row);

    // debugger;
    let prev = 0;
    let prevHeader = "";
    columnHeaders[i] = row.flatMap((cell, j) => {
      //   console.log(cell);
      if (cell._source === null) {
        return "";
      } else if (typeof cell.source() === "object") {
        const axisCellLength = cell._source._domain.length;
        if (i > 0 && typeof columnMatrix[i - 1][j]._source === "string") {
          //   console.log(j, columnHeaders[i - 1][j + 1]);
          let k = i;
          //   debugger;
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
  console.log(columnHeaders);
};

export const exportToExcel = (canvas) => {
  const data = extractPivotData(canvas);
  convertToExcel(data);
};
