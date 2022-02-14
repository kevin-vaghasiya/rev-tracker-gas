const getSideBar = () => {
  const template = HtmlService.createTemplateFromFile('sidebar');
  template.sheets = JSON.stringify(getSheets());
  SpreadsheetApp.getUi().showSidebar(
    template.evaluate().setTitle('Sheet Finder').setWidth(250)
  );
};

const getSheets = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  return sheets.map((sheet) => {
    const id = sheet.getSheetId();
    const name = sheet.getName();
    return { name, id };
  });
};

const setActiveSheet = (sheet_name: string) => {
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(sheet_name);
    if (!sheet) return;
    ss.setActiveSheet(sheet);
  } catch (error) {
    console.log(error);
  }
};
