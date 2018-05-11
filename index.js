import fs from "fs-extra";
import Excel from "exceljs";
import moment from "moment";
import _ from "lodash";
import path from "path";

import config from "./history/config";

// constant values use
const GSTCELL = `E40`;
const SUBTOTALCELL = `E39`;
const AMOUNTCELL = `E18`;
const TOTALCELL = `E41`;
const INVOICECELL = `E4`;
const DATESENTCELL = `E5`;
const DATEPERIODCELL = `A18`;

// folder path
const invoiceFolder = path.resolve(__dirname, "history");
// latest file path
const latestInvoice = path.resolve(
  __dirname,
  "history",
  `${config.latestInvoiceFileName}.xlsx`
);

// weekly wage input HERE!
const wage = 1000;

// Moment time default setting
moment.locale("en-AU");

const workbook = new Excel.Workbook();
workbook.xlsx.readFile(latestInvoice).then(function(wb) {
  // Copying latest sheet workbook sheet
  wb.properties.date1904 = true;
  // Search the list of wb worksheet list
  const filterdWorksheet = _.filter(wb._worksheets);
  // Locate the worksheet
  const latestWorksheet = wb.getWorksheet(
    filterdWorksheet[filterdWorksheet.length - 1].id
  );
  // Extract worksheet name
  const previousName = latestWorksheet.name;
  // Compose new worksheet name
  let week = parseInt(previousName.replace("W", ""));
  const newWorksheetName = `W${week + 1}`;

  // Change config file
  config.latestWeek = newWorksheetName;
  config.latestInvoiceFileName = `Invoice Fei ${config.latestWeek}`;
  // save config
  // fs.outputJson(`./history/config.json`, config, err => console.log(err))

  // Create a new workbook and worksheet
  const new_wb = new Excel.Workbook();

  new_wb.creator = "Fei";
  new_wb.lastModifiedBy = "Fei";
  new_wb.created = new Date();
  new_wb.modified = new Date();
  new_wb.lastPrinted = new Date();

  new_wb.properties.date1904 = true;
  const new_sheet = new_wb.addWorksheet(newWorksheetName);

  // Retrive associated invoices number and date information from previous sheet
  // Invoice Number
  const invoiceNumber = latestWorksheet.getCell(INVOICECELL).value;
  const newInvoiceNumber = `FA${parseInt(invoiceNumber.replace("FA", "")) + 1}`;

  // Invoice sent date
  const dateSent = latestWorksheet.getCell(DATESENTCELL).value;
  console.log(dateSent);
  const newDateSent = moment(dateSent, "DD-MMM-YY")
    .add(7, "days")
    .format("DD-MMM-YY");

  // Time about invoice includes
  const datePeriods = latestWorksheet.getCell(DATEPERIODCELL).value;
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
  new_sheet["_columns"] = _.cloneDeep(latestWorksheet["_columns"]);
  new_sheet["_rows"] = _.cloneDeep(latestWorksheet["_rows"]);

  // Change the values on the newly created sheet
  new_sheet.getCell(TOTALCELL).value = wage;
  new_sheet.getCell(GSTCELL).value = wage / 11;
  new_sheet.getCell(AMOUNTCELL).value = wage - new_sheet.getCell(GSTCELL).value;
  new_sheet.getCell(SUBTOTALCELL).value =
    wage - new_sheet.getCell(GSTCELL).value;
  // Change the invoice number
  new_sheet.getCell(INVOICECELL).value = newInvoiceNumber;
  // Change invoice sent date
  new_sheet.getCell(DATESENTCELL).value = newDateSent;
  // Change invoice periods
  new_sheet.getCell(DATEPERIODCELL).value = `${newStartDay} ~~ ${newEndDay}`;

  new_wb.xlsx
    .writeFile(
      path.resolve(__dirname, "history", `${config.latestInvoiceFileName}.xlsx`)
    )
    .then(() => {
        // save config
        fs.outputJson(`./history/config.json`, config, err => console.log(err))
    })
    .catch(function(err) {
      console.log(err);
    });
});
