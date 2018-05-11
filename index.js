import fs from 'fs-extra';
import Excel from 'exceljs';
import moment from 'moment';
import _ from 'lodash';
import path from 'path';

import config from './history/config.json';

// constant values use
const GSTCELL = 'E40';
const SUBTOTALCELL = 'E39';
const AMOUNTCELL = 'E18';
const TOTALCELL = 'E41';
const INVOICECELL = 'E4';
const DATESENTCELL = 'E5';
const DATEPERIODCELL = 'A18';

// latest file path
const latestInvoice = path.resolve(__dirname, 'history', `${config.latestInvoiceFileName}.xlsx`);

// weekly wage input HERE!
const wage = 1000;

// Moment time default setting
moment.locale('en-AU');

const workbook = new Excel.Workbook();

async function run() {
  const wb = await workbook.xlsx.readFile(latestInvoice);

  // Copying latest sheet workbook sheet
  wb.properties.date1904 = true;
  // Search the list of wb worksheet list
  const filterdWorksheet = _.filter(wb._worksheets);
  // Locate the worksheet
  const latestWorksheet = wb.getWorksheet(filterdWorksheet[filterdWorksheet.length - 1].id);
  // Extract worksheet name
  const previousName = latestWorksheet.name;
  // Compose new worksheet name
  const week = parseInt(previousName.replace('W', ''), 10);
  const newWorksheetName = `W${week + 1}`;

  // Change config file
  config.latestWeek = newWorksheetName;
  config.latestInvoiceFileName = `Invoice Fei ${config.latestWeek}`;
  // save config
  // fs.outputJson(`./history/config.json`, config, err => console.log(err))

  // Change config file
  config.latestWeek = newWorksheetName;
  config.latestInvoiceFileName = `Invoice Fei ${config.latestWeek}`;
  // save config
  // fs.outputJson(`./history/config.json`, config, err => console.log(err))

  // Create a new workbook and worksheet
  const newWorkbook = new Excel.Workbook();

  newWorkbook.creator = 'Fei';
  newWorkbook.lastModifiedBy = 'Fei';
  newWorkbook.created = new Date();
  newWorkbook.modified = new Date();
  newWorkbook.lastPrinted = new Date();

  newWorkbook.properties.date1904 = true;
  const newSheet = newWorkbook.addWorksheet(newWorksheetName);

  // Retrive associated invoices number and date information from previous sheet
  // Invoice Number
  const invoiceNumber = latestWorksheet.getCell(INVOICECELL).value;
  const newInvoiceNumber = `FA${parseInt(invoiceNumber.replace('FA', ''), 10) + 1}`;

  // Invoice sent date
  const dateSent = latestWorksheet.getCell(DATESENTCELL).value;
  const newDateSent = moment(dateSent, 'DD-MMM-YY')
    .add(7, 'days')
    .format('DD-MMM-YY');

  // Time about invoice includes
  const datePeriods = latestWorksheet.getCell(DATEPERIODCELL).value;
  const duration = datePeriods.split(' ~~ ');
  const startDay = duration[0];
  const endDay = duration[1];
  const newStartDay = moment(startDay, 'DD/MM/YYYY')
    .add(7, 'days')
    .format('DD/MM/YYYY');
  const newEndDay = moment(endDay, 'DD/MM/YYYY')
    .add(7, 'days')
    .format('DD/MM/YYYY');

  // Copy information from previous sheet to newly created sheet
  newSheet._columns = _.cloneDeep(latestWorksheet._columns);
  newSheet._rows = _.cloneDeep(latestWorksheet._rows);

  // Change the values on the newly created sheet
  newSheet.getCell(TOTALCELL).value = wage;
  newSheet.getCell(GSTCELL).value = wage / 11;
  newSheet.getCell(AMOUNTCELL).value = wage - newSheet.getCell(GSTCELL).value;
  newSheet.getCell(SUBTOTALCELL).value = wage - newSheet.getCell(GSTCELL).value;
  // Change the invoice number
  newSheet.getCell(INVOICECELL).value = newInvoiceNumber;
  // Change invoice sent date
  newSheet.getCell(DATESENTCELL).value = newDateSent;
  // Change invoice periods
  newSheet.getCell(DATEPERIODCELL).value = `${newStartDay} ~~ ${newEndDay}`;

  // Save file to history
  await newWorkbook.xlsx.writeFile(path.resolve(__dirname, 'history', `${config.latestInvoiceFileName}.xlsx`));
  // save config
  await fs.outputJson('./history/config.json', config);
  // Remove sendList folder item
  await fs.emptyDir(path.resolve(__dirname, 'sendList'));
  // Save it again to sendList folder
  await fs.copy(
    path.resolve(__dirname, 'history', `${config.latestInvoiceFileName}.xlsx`),
    path.resolve(__dirname, 'sendList', `${config.latestInvoiceFileName}.xlsx`),
  );
}

run();
