const xlsx = require("xlsx");
const fs = require("fs");
const process = require("process");

const argvFileName = process.argv[process.argv.length - 1]; //get the param (name file) from terminal
const xlsFile = xlsx.readFile(argvFileName);

const toSnakeCase = string => {
  const s = string.replace(/[^\w\s]/g, "").replace(/\s+/g, " ");
  return s
    .toLowerCase()
    .split(" ")
    .join("_");
};

const getFileJsonName = (originalFileName, sheetName) => {
  return `${originalFileName.split(".")[0]}.${sheetName.toLowerCase()}.json`;
};

const sheet2arr = (sheet, numberOfCol, numberOfRow) => {
  let rowArr;
  let rowNum = 0;
  let colNum = 0;
  const resultArr = [];
  const titleNameArr = [];
  const range = xlsx.utils.decode_range(sheet["!ref"]);

  for (colNum = range.s.c; colNum <= numberOfCol; colNum++) {
    rowArr = [];
    for (rowNum = range.s.r; rowNum <= numberOfRow; rowNum++) {
      const nextCell = sheet[xlsx.utils.encode_cell({ r: rowNum, c: colNum })];

      if (typeof nextCell !== "undefined") {
        // get array of title name then make the key in object
        if (colNum === 0) {
          titleNameArr.push(nextCell.w);
        } else {
          rowArr.push({ [toSnakeCase(titleNameArr[rowNum])]: nextCell.w });
        }
      }
    }
    // don't add empty row
    rowArr.length && resultArr.push(rowArr);
  }

  return resultArr;
};

xlsFile.SheetNames.map(sheetName => {
  const workbook = xlsFile.Sheets[sheetName];
  const XL_row_object = xlsx.utils.sheet_to_row_object_array(workbook);

  const result = sheet2arr(
    workbook,
    Object.keys(XL_row_object[0]).length + 2, //column --> ? dont know why missing 2 cols when using xlsx.utils.sheet_to_row_object_array
    XL_row_object.length + 1 //row ? dont know why missing 1 row when using xlsx.utils.sheet_to_row_object_array
  );

  fs.writeFileSync(
    getFileJsonName(argvFileName, sheetName),
    JSON.stringify(result)
  );
});
