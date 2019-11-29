const xlsx = require("xlsx");
const fs = require("fs");
const process = require("process");

const argvFileName = process.argv[process.argv.length - 1];
const xlsFile = xlsx.readFile(argvFileName);

function toSnakeCase(string) {
  const s = string.replace(/[^\w\s]/g, "").replace(/\s+/g, " ");
  return s
    .toLowerCase()
    .split(" ")
    .join("_");
}

const sheet2arr = (sheet, numberOfCol, numberOfRow) => {
  var result = [];
  var row;
  var rowNum;
  var colNum;
  var range = xlsx.utils.decode_range(sheet["!ref"]);
  var titleName = [];
  for (rowNum = range.s.r; rowNum < numberOfCol; rowNum++) {
    row = [];
    for (colNum = range.s.c; colNum < numberOfRow; colNum++) {
      var nextCell = sheet[xlsx.utils.encode_cell({ r: colNum, c: rowNum })];

      if (typeof nextCell !== "undefined") {
        rowNum === 0 && titleName.push(nextCell.w);
        row.push({ [toSnakeCase(titleName[colNum])]: nextCell.w });
      } else {
        row.push({});
      }
    }
    if (row.length) {
      result.push(row);
    }
  }

  fs.writeFileSync(`${argvFileName}.json`, JSON.stringify(result), () => {
    console.log(`The ${argvFileName}.json is created`);
  });
  return result;
};

xlsFile.SheetNames.forEach(function(sheetName) {
  var XL_row_object = xlsx.utils.sheet_to_row_object_array(
    xlsFile.Sheets[sheetName]
  );
  sheet2arr(
    xlsFile.Sheets[sheetName],
    Object.keys(XL_row_object[0]).length, //column
    XL_row_object.length //row
  );
});
