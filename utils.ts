const onOpen = () => {
  SpreadsheetApp.getUi()
    .createMenu('Menu')
    .addItem('Authorize', 'authorize')
    .addItem('Test  Script', 'calculateRevenue')
    .addItem('Clear Cache', 'clearCache')
    .addItem('Start Script', 'startScript')
    .addItem('Stop Script', 'stopScript')
    .addToUi();
};

const clearCache = () => {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.CACHE);
  sheet.clear();
};

const authorize = () => {
  Logger.log('Start');
};

const startScript = () => {
  stopScript();
  ScriptApp.newTrigger('calculateRevenue').timeBased().everyDays(21).create();
  SpreadsheetApp.getUi().alert('Script Started Successfully.');
};

const stopScript = () => {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
};

const addError = (error: string) => {
  try {
    SpreadsheetApp.getActive()
      .getSheetByName(SHEET_NAMES.ERROR_LOGS)
      .appendRow([new Date(), error]);
  } catch (e) {
    console.log(e);
  }
};

const getMonthYearFromDate = (dt: number | string) => {
  try {
    const dateObj = new Date(dt);
    return { month: dateObj.getMonth(), year: dateObj.getFullYear() };
  } catch (error) {
    return {};
  }
};

const getFormattedDate = (dt: number | string) => {
  try {
    if (!dt) return '';
    if (!isNaN(Number(dt))) dt = Number(dt);
    const dateObj = new Date(dt);
    const date = dateObj.getDate();
    const month = dateObj.getMonth();
    const year = dateObj.getFullYear();
    return `${month < 9 ? '0' + (month + 1) : month + 1}/${
      date < 10 ? '0' + date : date
    }/${year}`;
  } catch (error) {
    return '';
  }
};

const getLeadFromLeadId = (id: string, data: string[][], idIndex: number) => {
  for (let i = 1; i < data.length; i++) {
    const lead_id = data[i][idIndex];
    if (id !== lead_id) continue;
    return { row_num: i + 1, row: data[i] };
  }
  return { row_num: -1, row: [] };
};

const columnToLetter = (column: number) => {
  let temp,
    letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
};

const getDeductionFromSaleType = (sale_type: string) => {
  if (sale_type == SALE_TYPES.CO_OP_LISTING)
    return AUTOMATIC_DEDUCTIONS.CO_OP_QW;
  if (sale_type == SALE_TYPES.QW_LISTING_CO_OP)
    return AUTOMATIC_DEDUCTIONS.QW_CO_OP;
  if (sale_type == SALE_TYPES.QW_LISTING_QW) return AUTOMATIC_DEDUCTIONS.QW_QW;
  return 0;
};
