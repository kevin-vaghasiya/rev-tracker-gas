const getHeadersAndData = (ss: GoogleAppsScript.Spreadsheet.Spreadsheet) => {
  const formSheet = ss.getSheetByName(QWBT_SHEET_NAMES.FORM_RESPONSES);
  const data = formSheet
    .getRange(1, 1, formSheet.getLastRow(), formSheet.getLastColumn())
    .getValues();
  const header: string[] = data[0];
  return { header, data };
};

const getQwHeaderIndexes = (
  ss: GoogleAppsScript.Spreadsheet.Spreadsheet,
  header: string[]
): IHeaderIndexes => {
  const name_mapping_sheet = ss.getSheetByName(QWBT_SHEET_NAMES.NAME_MAPPING);
  const nameData = name_mapping_sheet
    .getRange(
      1,
      1,
      name_mapping_sheet.getLastRow(),
      name_mapping_sheet.getLastColumn()
    )
    .getValues();
  const mappings = {};
  for (let i = 0; i < header.length; i++) {
    const headerName = header[i];
    if (!headerName) continue;
    let ifAdded = false;
    for (let j = 1; j < nameData.length; j++) {
      const [name, shortName] = nameData[j];
      if (!name) continue;
      if (headerName == name && shortName) {
        mappings[shortName] = { index: i };
        ifAdded = true;
        break;
      }
    }
    if (!ifAdded) mappings[headerName] = { index: i };
  }
  return mappings;
};

const getChangeLogIds = (ss: GoogleAppsScript.Spreadsheet.Spreadsheet) => {
  try {
    const changeLogSheet = ss.getSheetByName(QWBT_SHEET_NAMES.CHANGE_LOGS);
    const data = changeLogSheet
      .getRange(1, 1, changeLogSheet.getLastRow(), 1)
      .getValues();
    const ids = [];
    for (let i = 0; i < data.length; i++) {
      const [lead_id] = data[i];
      if (!lead_id) continue;
      if (ids.indexOf(lead_id) != -1) continue;
      ids.push(lead_id);
    }
    changeLogSheet.clear(); //TODO
    return ids;
  } catch (error) {
    return [];
  }
};
