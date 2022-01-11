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
