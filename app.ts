const calculateRevenue = () => {
  const ss = SpreadsheetApp.getActive();
  try {
    const { year, qw_sheet_url } = getConfigData(ss);
    const leads: IQWLead[] = getQWLeads(qw_sheet_url, year);
    if (!leads.length) return;

    const nameSheetCache: ISheetCache = {};
    const headersCache: IHeadersCache = {};
    const { cache, cacheSheet } = getCacheSheetAndData(ss);
    upsertLeads(
      ss,
      leads,
      year,
      cache,
      cacheSheet,
      nameSheetCache,
      headersCache
    );
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
  nameSheetCache: ISheetCache,
  headersCache: IHeadersCache
) => {
  const listingSheet = ss.getSheetByName(SHEET_NAMES.LISTINGS);
  for (let i = 0; i < leads.length; i++) {
    const lead = leads[i];
    const { Id } = lead;
    const { month, year } = getMonthYearFromDate(lead['Settlement Date']);
    if (cache[Id]) {
      const {
        sheet_name,
        date: { month: prevMonth },
      } = cache[Id];
      if (year !== currentYear) {
        const sheet = getSheetFromName(ss, sheet_name, nameSheetCache);
        deleteLeadFromSheetAndCache(sheet, Id, cache[Id].index, cacheSheet);
        continue;
      } else if (month == prevMonth) {
        //TODO update logic
      } else {
        // Change Month
      }
    } else if (year == currentYear) {
      const headers = getHeaderFromCache(
        ss,
        SHEET_NAMES.LISTINGS,
        headersCache,
        nameSheetCache
      );

      const row = new Array(34);
      addCalculatedFields(lead);
      for (const key in headers) {
        if (Object.prototype.hasOwnProperty.call(headers, key)) {
          const header = headers[key];
          if (header && lead[key]) row[header.index] = lead[key];
        }
      }
      //Add Projected Rev Fields
      listingSheet.appendRow(row);
      addLeadToCache(cacheSheet, {
        id: Id,
        date: { year, month },
        sheet_name: SHEET_NAMES.LISTINGS,
        type: lead['Sale Type'],
      });
    }
  }
};

const addCalculatedFields = (lead: IQWLead) => {
  const projectedRev =
    ((Number(lead['Sale Price']) * Number(lead['Commission Rate'])) / 100) *
    0.65; // AGENT COMMISSION
  lead[KEY_NAMES.PROJECTED_REV] = projectedRev;
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
  const data = cacheSheet.getDataRange().getValues();
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

const getQWLeads = (url: string, year: number) => {
  const qwSs = SpreadsheetApp.openByUrl(url);
  const { header, data } = getHeadersAndData(qwSs);
  const indexes = getQwHeaderIndexes(qwSs, header);
  const change_log_ids = getChangeLogIds(qwSs);
  const leads = [];

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
        const commission = element[indexes[QW_KEY_NAMES.COMMISSION_RATE].index];
        const doc1_link = element[indexes[QW_KEY_NAMES.DOC1].index];
        const mgt_notes = element[indexes[QW_KEY_NAMES.MGT_NOTES].index];
        const mls_no = element[indexes[QW_KEY_NAMES.MLS_NO].index];
        const property_address =
          element[indexes[QW_KEY_NAMES.PROPERTY_ADDRESS].index];
        const sale_price = element[indexes[QW_KEY_NAMES.SALE_PRICE].index];
        const timestamp = element[indexes[QW_KEY_NAMES.TIMESTAMP].index];
        const aos_date = element[indexes[QW_KEY_NAMES.AOS_DATE].index];
        const agent_2 = element[indexes[QW_KEY_NAMES.QW_AGENT_2].index];
        leads.push({
          [KEY_NAMES.ID]: id,
          [KEY_NAMES.AGENT_NAME]: agent_name,
          [KEY_NAMES.COMMISSION_RATE]: commission,
          [KEY_NAMES.DOC1]: doc1_link,
          [KEY_NAMES.MGT_NOTES]: mgt_notes,
          [KEY_NAMES.MLS_NO]: mls_no,
          [KEY_NAMES.PROPERTY_ADDRESS]: property_address,
          [KEY_NAMES.SALE_PRICE]: sale_price,
          [KEY_NAMES.SALE_TYPE]: sale_type,
          [KEY_NAMES.TIMESTAMP]: timestamp,
          [KEY_NAMES.SETTLEMENT_DATE]: s_date,
          [KEY_NAMES.AOS_DATE]: aos_date,
          [KEY_NAMES.QW_AGENT_2]: agent_2,
        });
      }
    } catch (error) {
      console.log(error);
    }
  }
  return leads;
};

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

const getHeaderFromCache = (
  ss: GoogleAppsScript.Spreadsheet.Spreadsheet,
  sheet_name: string,
  headerCache: IHeadersCache,
  cacheSheet: ISheetCache
): IHeaderIndexes => {
  try {
    if (headerCache[sheet_name]) return headerCache[sheet_name];
    return getHeaderIndexes(getSheetFromName(ss, sheet_name, cacheSheet));
  } catch (error) {
    return {};
  }
};

const getHeaderIndexes = (
  sheet: GoogleAppsScript.Spreadsheet.Sheet
): IHeaderIndexes => {
  const [header] = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
  const mappings = {};
  for (let i = 0; i < header.length; i++) {
    const headerName = header[i];
    if (!headerName) continue;
    mappings[headerName] = { index: i };
  }
  return mappings;
};
