const calculateRevenue = () => {
  const ss = SpreadsheetApp.getActive();
  try {
    const { year, qw_sheet_url } = getConfigData(ss);
    const leads: IQWLead[] = getQWLeads(qw_sheet_url, year);
    if (!leads.length) return;
    const nameSheetCache: ISheetCache = {};
    const { cache, cacheSheet } = getCacheSheetAndData(ss);
    upsertLeads(ss, leads, year, cache, cacheSheet, nameSheetCache);
  } catch (error) {
    SpreadsheetApp.getUi().alert(error);
  }
};

const upsertLeads = (
  ss: GoogleAppsScript.Spreadsheet.Spreadsheet,
  leads: IQWLead[],
  currentYear: number,
  cache: ICache,
  cacheSheet: GoogleAppsScript.Spreadsheet.Sheet,
  nameSheetCache: ISheetCache
) => {
  const listingSheet = ss.getSheetByName(SHEET_NAMES.LISTINGS);
  for (let i = 0; i < leads.length; i++) {
    const {
      id,
      sale_type,
      timestamp,
      property_address,
      agent_name,
      sale_price,
      commission,
      mgt_notes,
      doc1_link,
      mls_no,
      settlement_date,
    } = leads[i];

    // if cache data exist
    const { month, year } = getMonthYearFromDate(settlement_date);
    if (cache[id]) {
      const { sheet_name, type } = cache[id];
      //   console.log({ year, currentYear });
      if (year !== currentYear) {
        const sheet = getSheetFromName(ss, sheet_name, nameSheetCache);
        deleteLeadFromSheetAndCache(sheet, id, cache[id].index, cacheSheet);
        continue;
      } else {
        //TODO update logic
      }
    } else if (year == currentYear) {
      listingSheet.appendRow([
        timestamp,
        sale_type,
        property_address,
        agent_name,
        sale_price,
        commission,
        ,
        mgt_notes,
        doc1_link,
        mls_no,
        id,
      ]);
      addLeadToCache(cacheSheet, {
        id,
        date: { year, month },
        sheet_name: SHEET_NAMES.LISTINGS,
        type: sale_type,
      });
    }
  }
};

const addLeadToCache = (
  cacheSheet: GoogleAppsScript.Spreadsheet.Sheet,
  data: {
    id: string;
    date: {
      month: number;
      year: number;
    };
    sheet_name: string;
    type: string;
  }
) => {
  cacheSheet.appendRow([data.id, JSON.stringify(data)]);
};

const deleteLeadFromSheetAndCache = (
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  id: string,
  index: number,
  cacheSheet: GoogleAppsScript.Spreadsheet.Sheet
) => {
  const data = sheet
    .getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn())
    .getValues();
  for (let i = 0; i < data.length; i++) {
    const [timestamp] = data[i];
    if (!timestamp) continue;
    const did = data[i][10];
    if (did != id) continue;
    sheet.deleteRow(i + 1);
    break;
  }
  cacheSheet.deleteRow(index + 1);
};

const getCacheSheetAndData = (ss: GoogleAppsScript.Spreadsheet.Spreadsheet) => {
  const cacheSheet = ss.getSheetByName(SHEET_NAMES.CACHE);
  const data = cacheSheet
    .getRange(1, 1, cacheSheet.getLastRow(), cacheSheet.getLastColumn())
    .getValues();
  const cache: ICache = {};
  for (let i = 0; i < data.length; i++) {
    const [id, str] = data[i];
    if (!id) continue;
    const cacheData = JSON.parse(str) as {
      date: {
        month: number;
        year: number;
      };
      sheet_name: string;
      type: string;
    };
    cache[id] = { ...cacheData, index: i };
  }
  return { cacheSheet, cache };
};

const getQWLeads = (url: string, year: number): IQWLead[] => {
  const qwSs = SpreadsheetApp.openByUrl(url);
  const { header, data } = getHeadersAndData(qwSs);
  const indexes = getHeaderIndexes(qwSs, header);
  const change_log_ids = getChangeLogIds(qwSs);
  const leads: IQWLead[] = [];

  for (let i = 1; i < data.length; i++) {
    try {
      const element = data[i];
      if (!element[0]) continue;
      const id = element[indexes[QW_KEY_NAMES.ID].index];
      const sale_type = element[indexes[QW_KEY_NAMES.SALE_TYPE].index];
      let s_date = null;
      if (sale_type == SALE_TYPES.QW_LISTING_ACTIVE)
        s_date = element[indexes[QW_KEY_NAMES.TIMESTAMP].index];
      else s_date = element[indexes[QW_KEY_NAMES.SETTLEMENT_DATE].index];
      const s_year = new Date(s_date).getFullYear();

      if (
        (id && change_log_ids.indexOf(id) !== -1) ||
        (s_date && year == s_year)
      ) {
        const agent_name = element[indexes[QW_KEY_NAMES.AGENT_NAME].index];
        const commission = element[indexes[QW_KEY_NAMES.COMMISSION_TS].index];
        const doc1_link = element[indexes[QW_KEY_NAMES.DOC1].index];
        const mgt_notes = element[indexes[QW_KEY_NAMES.MGT_NOTES].index];
        const mls_no = element[indexes[QW_KEY_NAMES.MLS_NO].index];
        const property_address =
          element[indexes[QW_KEY_NAMES.PROPERTY_ADDRESS].index];
        const sale_price = element[indexes[QW_KEY_NAMES.SALE_PRICE].index];
        const timestamp = element[indexes[QW_KEY_NAMES.TIMESTAMP].index];
        leads.push({
          id,
          agent_name,
          commission,
          doc1_link,
          mgt_notes,
          mls_no,
          property_address,
          sale_price,
          sale_type,
          timestamp,
          settlement_date: s_date,
        });
      }
    } catch (error) {
      console.log(error);
    }
  }
  return leads;
};

const deleteCache = () => {};

const getConfigData = (ss: GoogleAppsScript.Spreadsheet.Spreadsheet) => {
  const configSheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  const data = configSheet.getRange('A1:B2').getValues();
  const year = data[0][1];
  const qw_sheet_url = data[1][1];
  if (!year) throw new Error('Year Not Found in ' + SHEET_NAMES.CONFIG);
  if (!qw_sheet_url)
    throw new Error('QW Sheet Url Not Found in ' + SHEET_NAMES.CONFIG);
  return { year: Number(year), qw_sheet_url };
};

const getSheetFromName = (
  ss: GoogleAppsScript.Spreadsheet.Spreadsheet,
  name: string,
  sheetCache: ISheetCache
) => {
  if (sheetCache[name]) return sheetCache[name];
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  sheetCache[name] = sheet;
  return sheet;
};
