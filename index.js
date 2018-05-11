const Excel = require("exceljs");
const moment = require("moment");
const _ = require("lodash");
const path = require("path");

// constant values use
const GSTCELL = `E40`;
const SUBTOTALCELL = `E39`;
const AMOUNTCELL = `E18`;
const TOTALCELL = `E41`;
const INVOICECELL = `E4`;
const DATESENTCELL = `E5`;
const DATEPERIODCELL = `A18`;

// Paths
const xlsxReadOriginalPath = path.resolve(__dirname, "test/fei.xlsx");
const xlsxReadFilePath = path.resolve(__dirname, "test/readTest001.xlsx");
const xlsxWriteFilePath = path.resolve(__dirname, "test/writeTest001.xlsx");
// File extracted
// const xlsxInvoice = path.resolve(__dirname, "test/Invoice.xlsx");
const xlsxInvoice = path.resolve(__dirname, "history/Invoice.xlsx");

// INPUT VALUES GOES HERE!!!
const salary = 2282.5;

// Other setting
moment.locale("en-AU");

const workbook = new Excel.Workbook();
workbook.xlsx.readFile(xlsxInvoice).then(function(wb) {
  //Create new sheet using latest sheet name
  wb.properties.date1904 = true;
  const filterdWorksheet = _.filter(wb._worksheets);
  //   console.log(filterdWorksheet)
  const previouseWorksheet = wb.getWorksheet(
    filterdWorksheet[filterdWorksheet.length - 1].id
  );
  const previousName = previouseWorksheet.name;
  let week = parseInt(previousName.replace("W", ""));
  const newWorksheetName = `W${week + 1}`;
  //   console.log(wb);
  let sheet = wb.addWorksheet(newWorksheetName, {
    views: [{ xSplit: 1, ySplit: 1 }]
  });

  // Create a new workbook
  // const new_wb = new Excel.Workbook();

  // new_wb.creator = "Fei";
  // new_wb.lastModifiedBy = "Fei";
  // new_wb.created = new Date();
  // new_wb.modified = new Date();
  // new_wb.lastPrinted = new Date();

  // new_wb.properties.date1904 = true;
  // const new_sheet = workbook.addWorksheet(newWorksheetName);

  // Debug
  console.log(
    `Previous worksheet id: ${previouseWorksheet.id}, name: ${
      previouseWorksheet.name
    }`
  );
  console.log(`current worksheet id: ${sheet.id}, name: ${sheet.name}`);

  // Retrive associated invoices number and date information from previous sheet
  // Invoice Number
  const invoiceNumber = previouseWorksheet.getCell(INVOICECELL).value;
  const newInvoiceNumber = `FA${parseInt(invoiceNumber.replace("FA", "")) + 1}`;

  // Invoice sent date
  const dateSent = previouseWorksheet.getCell(DATESENTCELL).value;
  console.log(dateSent)
  const newDateSent = moment(dateSent, "DD-MMM-YY")
    .add(7, "days")
    .format("DD-MMM-YY");

  //
  //   console.log(dateSent);
  //   console.log(newDateSent);
  // Time about invoice includes
  const datePeriods = previouseWorksheet.getCell(DATEPERIODCELL).value;
  const duration = datePeriods.split(" ~~ ");
  const startDay = duration[0];
  const endDay = duration[1];
  const newStartDay = moment(startDay, "DD/MM/YYYY")
    .add(7, "days")
    .format("DD/MM/YYYY");
  const newEndDay = moment(endDay, "DD/MM/YYYY")
    .add(7, "days")
    .format("DD/MM/YYYY");

  //Copy information from previous sheet to newly created sheet
  sheet["_columns"] = _.cloneDeep(previouseWorksheet["_columns"]);
  sheet["_rows"] = _.cloneDeep(previouseWorksheet["_rows"]);
  //   sheet["_columns"] = Object.assign({}, previouseWorksheet["_columns"]);
  //   sheet["_rows"] = Object.assign({}, previouseWorksheet["_rows"]);

  // Change the values on the newly created sheet
  sheet.getCell(TOTALCELL).value = salary;
  sheet.getCell(GSTCELL).value = salary / 11;
  sheet.getCell(AMOUNTCELL).value = salary - sheet.getCell(GSTCELL).value;
  sheet.getCell(SUBTOTALCELL).value = salary - sheet.getCell(GSTCELL).value;
  // Change the invoice number
  sheet.getCell(INVOICECELL).value = newInvoiceNumber;
  // Change invoice sent date
  sheet.getCell(DATESENTCELL).value = newDateSent;
  // Change invoice periods
  sheet.getCell(DATEPERIODCELL).value = `${newStartDay} ~~ ${newEndDay}`;

  // Save changed onto local file
  wb.xlsx
    .writeFile(xlsxInvoice)
    .then(function(err) {
      console.log(err);
    })
    .catch(function(err) {
      console.log(err);
    });
  // Save changes into new workbook file
  // new_wb.writeFile;
});
