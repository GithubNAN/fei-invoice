const path = require("path");

const xlsxReadFilePath = path.resolve(__dirname, "test/readTest001.xlsx");
const xlsxWriteFilePath = path.resolve(__dirname, "test/writeTest001.xlsx");

const Excel = require("exceljs");

const workbook = new Excel.Workbook();

workbook.creator = "Test";
workbook.lastModifiedBy = "him";
workbook.created = new Date(1985, 8, 30);
workbook.modified = new Date();
workbook.lastPrinted = new Date(2017, 9, 9);

workbook.properties.date1904 = true;

workbook.views = [
  {
    x: 0,
    y: 0,
    width: 10000,
    height: 20000,
    firstSheet: 0,
    activeTab: 1,
    visibility: "visible"
  }
];

const sheet = workbook.addWorksheet("test sheet", {
  properties: {
    tabColor: {
      argb: "FFC0000"
    },
    showGridLines: false
  }
});

workbook.xlsx.writeFile(xlsxWriteFilePath)