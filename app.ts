const calculateRevenue = () => {
  const ss = SpreadsheetApp.getActive();
  try {
    const { year, qw_sheet_url } = getConfigData(ss);
    const leads: IQWLead[] = getQWLeads(qw_sheet_url, year);
    if (!leads.length) return;

    const nameSheetCache: ISheetCache = {};
    const headersCache: IHeadersCache = {};
    const { cache, cacheSheet } = getCacheSheetAndData(ss);
    const agents_data = getAgentsData(ss);
    upsertLeads(
      ss,
      leads,
      year,
      cache,
      cacheSheet,
      nameSheetCache,
      headersCache,
      agents_data
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert(error);
  }
};

const upsertLeads = (
  ss: GoogleAppsScript.Spreadsheet.Spreadsheet,
  leads: IQWLead[],
  currentYear: number,
  cache: ICacheData,
  cacheSheet: GoogleAppsScript.Spreadsheet.Sheet,
  nameSheetCache: ISheetCache,
  headersCache: IHeadersCache,
  agents_data: IAgentsData
) => {
  for (let i = 0; i < leads.length; i++) {
    const lead = leads[i];
    const { Id, 'Sale Type': sale_type } = lead;
    const { month, year } = getMonthYearFromDate(lead['Settlement Date']);
    const current_sheet_name = getSheetNameFromLead(sale_type, month);

    if (cache[Id]) {
      const { sheet_name } = cache[Id];
      // when year is changed => delete lead
      if (year !== currentYear || !current_sheet_name) {
        const sheet = getSheetFromName(ss, sheet_name, nameSheetCache);
        const headers = getHeaderFromCache(
          ss,
          current_sheet_name,
          headersCache,
          nameSheetCache
        );
        deleteLeadFromSheet(sheet, Id, headers);
        deleteLeadFromCache(cache[Id].index, cacheSheet);
      } else if (current_sheet_name !== sheet_name) {
        // when sheet is changed due to data change
        updateLeadInDifferentSheet(
          ss,
          sheet_name,
          current_sheet_name,
          lead,
          cache[Id],
          cacheSheet,
          nameSheetCache,
          headersCache,
          agents_data
        );
      } else {
        // when sheet is same
        updateLeadInSameSheet(
          ss,
          sheet_name,
          lead,
          nameSheetCache,
          headersCache,
          agents_data
        );
      }
    } else if (year == currentYear) {
      addLeadToSheet(
        ss,
        current_sheet_name,
        lead,
        nameSheetCache,
        headersCache,
        agents_data
      );
      addLeadToCache(cacheSheet, {
        id: Id,
        date: { year, month },
        sheet_name: current_sheet_name,
        type: sale_type,
      });
    }
  }
};

const addLeadToSheet = (
  ss: GoogleAppsScript.Spreadsheet.Spreadsheet,
  current_sheet_name: string,
  lead: IQWLead,
  nameSheetCache: ISheetCache,
  headersCache: IHeadersCache,
  agentsData: IAgentsData
) => {
  const sheet = getSheetFromName(ss, current_sheet_name, nameSheetCache);
  const headers = getHeaderFromCache(
    ss,
    current_sheet_name,
    headersCache,
    nameSheetCache
  );
  const row = new Array(34);
  addCalculatedFields(lead, sheet.getLastRow() + 1, headers, agentsData);
  for (const key in headers) {
    if (Object.prototype.hasOwnProperty.call(headers, key)) {
      const header = headers[key];
      if (header && lead[key]) row[header.index] = lead[key];
    }
  }
  sheet.appendRow(row);
};

const updateLeadInDifferentSheet = (
  ss: GoogleAppsScript.Spreadsheet.Spreadsheet,
  sheet_name: string,
  current_sheet_name: string,
  lead: IQWLead,
  cache: ICache,
  cacheSheet: GoogleAppsScript.Spreadsheet.Sheet,
  nameSheetCache: ISheetCache,
  headersCache: IHeadersCache,
  agentsData: IAgentsData
) => {
  const sheet = getSheetFromName(ss, sheet_name, nameSheetCache);
  const headers = getHeaderFromCache(
    ss,
    sheet_name,
    headersCache,
    nameSheetCache
  );
  const { row, row_num } = getRowNumberFromLead(sheet, lead, headers);
  for (let i = 0; i < MANUAL_FIELDS.length; i++) {
    const field = MANUAL_FIELDS[i];
    const index = headers[field]?.index;
    if (index == 0 || index) {
      if (!row[index]) continue;
      lead[field] = row[index];
    }
  }
  addCalculatedFields(lead, row_num, headers, agentsData);
  addLeadToSheet(
    ss,
    current_sheet_name,
    lead,
    nameSheetCache,
    headersCache,
    agentsData
  );
  const id = lead['Id'];
  deleteLeadFromSheet(sheet, id, headers);
  const { month, year } = getMonthYearFromDate(lead['Settlement Date']);
  updateCacheInSheet(
    cache,
    cacheSheet,
    JSON.stringify({
      id,
      date: { year, month },
      sheet_name: current_sheet_name,
      type: lead['Sale Type'],
    })
  );
};

const updateLeadInSameSheet = (
  ss: GoogleAppsScript.Spreadsheet.Spreadsheet,
  current_sheet_name: string,
  lead: IQWLead,
  nameSheetCache: ISheetCache,
  headersCache: IHeadersCache,
  agentsData: IAgentsData
) => {
  const sheet = getSheetFromName(ss, current_sheet_name, nameSheetCache);
  const headers = getHeaderFromCache(
    ss,
    current_sheet_name,
    headersCache,
    nameSheetCache
  );
  const { row, row_num } = getRowNumberFromLead(sheet, lead, headers);
  addCalculatedFields(lead, row_num, headers, agentsData);
  if (row_num == -1) return;
  for (const key in headers) {
    if (Object.prototype.hasOwnProperty.call(headers, key)) {
      const header = headers[key];
      if (!header) continue;
      if ((lead[key] == 0 || lead[key]) && row[header.index] != lead[key]) {
        sheet.getRange(row_num, header.index + 1).setValue(lead[key]);
      }
    }
  }
};

const getRowNumberFromLead = (
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  lead: IQWLead,
  headers: IHeaderIndexes
) => {
  const data = sheet.getDataRange().getValues();
  return getLeadFromLeadId(lead['Id'], data, headers['Id'].index);
};

const addCalculatedFields = (
  lead: IQWLead,
  row_num: number,
  headers: IHeaderIndexes,
  agentsData: IAgentsData
) => {
  const sale_type = lead['Sale Type'];
  const agent_commission = calculateAgentsCommission(
    lead['Agent Name'],
    lead['AoS Date'],
    agentsData
  );

  if (sale_type == SALE_TYPES.QW_LISTING_ACTIVE) {
    const projectedRev =
      ((Number(lead['Sale Price']) * Number(lead['Commission Rate'])) / 100) *
      agent_commission; // AGENT COMMISSION
    lead[KEY_NAMES.PROJECTED_REV] = projectedRev;
  } else {
    const manual_commission_index = headers[KEY_NAMES.MANUAL_COMMISSION].index;
    const manual_commission_column =
      columnToLetter(manual_commission_index + 1) + `${row_num}`;

    const rev =
      (Number(lead['Sale Price']) * Number(lead['Commission Rate'])) / 100;
    const commission = `=IF(${manual_commission_column}>0,0,${rev})`;
    lead[KEY_NAMES.CALCULATED_COMMISSION] = commission;

    const manual_deduction_index = headers[KEY_NAMES.MANUAL_DEDUCTION].index;
    const manual_deduction_column =
      columnToLetter(manual_deduction_index + 1) + `${row_num}`;
    const deduction = getDeductionFromSaleType(sale_type);

    lead[
      KEY_NAMES.AUTOMATIC_DEDUCTION
    ] = `=IF(${manual_deduction_column}>0,0,${deduction})`;

    lead[
      KEY_NAMES.REVENUE_FOR_SPLIT
    ] = `=IF(${manual_commission_column}>0,${manual_commission_column},${rev})-IF(${manual_deduction_column}>0,${manual_deduction_column},${deduction})`;

    const revenue_index = headers[KEY_NAMES.REVENUE_FOR_SPLIT].index;
    const revenue_column = columnToLetter(revenue_index + 1) + `${row_num}`;
    lead[KEY_NAMES.AGENT_REVENUE] = `=${revenue_column}*${agent_commission}`;
    lead[KEY_NAMES.QW_REVENUE] = `=${revenue_column}*1/${agent_commission}`;

    // sale type QW | QW
    if (sale_type == SALE_TYPES.QW_LISTING_QW) {
      const manual_commission_index_2 =
        headers[KEY_NAMES.MANUAL_COMMISSION_2].index;
      const manual_commission_column_2 =
        columnToLetter(manual_commission_index_2 + 1) + `${row_num}`;

      const rev =
        (Number(lead['Sale Price']) * Number(lead['Commission Rate'])) / 100;
      const commission = `=IF(${manual_commission_column_2}>0,0,${rev})`;
      lead[KEY_NAMES.CALCULATED_COMMISSION_2] = commission;

      const manual_deduction_index_2 =
        headers[KEY_NAMES.MANUAL_DEDUCTION_2].index;
      const manual_deduction_column_2 =
        columnToLetter(manual_deduction_index_2 + 1) + `${row_num}`;
      const deduction = getDeductionFromSaleType(sale_type);

      lead[
        KEY_NAMES.AUTOMATIC_DEDUCTION_2
      ] = `=IF(${manual_deduction_column_2}>0,0,${deduction})`;

      lead[
        KEY_NAMES.REVENUE_FOR_SPLIT_2
      ] = `=IF(${manual_commission_column_2}>0,${manual_commission_column_2},${rev})-IF(${manual_deduction_column_2}>0,${manual_deduction_column_2},${deduction})`;

      const agent2Commission = calculateAgentsCommission(
        lead['QWAgent2'],
        lead['AoS Date'],
        agentsData
      );
      const revenue_2_index = headers[KEY_NAMES.REVENUE_FOR_SPLIT_2].index;
      const revenue_2_column =
        columnToLetter(revenue_2_index + 1) + `${row_num}`;
      lead[
        KEY_NAMES.AGENT_REVENUE_2
      ] = `=${revenue_2_column}*${agent2Commission}`;
      lead[
        KEY_NAMES.QW_REVENUE_2
      ] = `=${revenue_2_column}*1/${agent2Commission}`;
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

const updateLeadCacheInSheet = (
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

const deleteLeadFromSheet = (
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  id: string,
  headers: IHeaderIndexes
) => {
  const data = sheet
    .getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn())
    .getValues();
  const idIndex = headers[KEY_NAMES.ID]?.index;
  for (let i = 0; i < data.length; i++) {
    const [timestamp] = data[i];
    if (!timestamp) continue;
    const did = data[i][idIndex];
    if (did != id) continue;
    sheet.deleteRow(i + 1);
    break;
  }
};

const deleteLeadFromCache = (
  index: number,
  cacheSheet: GoogleAppsScript.Spreadsheet.Sheet
) => {
  cacheSheet.deleteRow(index + 1);
};

const updateCacheInSheet = (
  cache: ICache,
  cacheSheet: GoogleAppsScript.Spreadsheet.Sheet,
  data: string
) => {
  cacheSheet.getRange(cache.index + 1, 2).setValue(data);
};

const getCacheSheetAndData = (ss: GoogleAppsScript.Spreadsheet.Spreadsheet) => {
  const cacheSheet = ss.getSheetByName(SHEET_NAMES.CACHE);
  const data = cacheSheet.getDataRange().getValues();
  const cache: ICacheData = {};
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
        (s_date && year == s_year) //TODO only add specific sale types
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
        const commission_2 = element[indexes[QW_KEY_NAMES.COMMISSION_2].index];

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
          [KEY_NAMES.COMMISSION_2]: commission_2,
        });
      }
    } catch (error) {
      console.log(error);
    }
  }
  return leads;
};

const getAgentsData = (ss: GoogleAppsScript.Spreadsheet.Spreadsheet) => {
  const splitSheet = ss.getSheetByName(SHEET_NAMES.SPLITS);
  const split_data = splitSheet
    .getRange(1, 1, splitSheet.getLastRow(), splitSheet.getLastColumn())
    .getValues();
  const agents_data: IAgentsData = {};
  for (let i = 0; i < split_data.length; i++) {
    const [fun, fin, l_name, email, q1, q2, q3, q4] = split_data[i];
    if (!l_name) continue;
    agents_data[l_name] = {
      q1: Number(q1) || 0,
      q2: Number(q2) || 0,
      q3: Number(q3) || 0,
      q4: Number(q4) || 0,
    };
  }
  return agents_data;
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
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name.includes('Listings')) sheet.appendRow(HEADERS.LISTING);
    else if (name.includes('QW|Co-Op')) sheet.appendRow(HEADERS.QW_CO_OP);
    else if (name.includes('Co-Op|QW')) sheet.appendRow(HEADERS.CO_OP_QW);
    else if (name.includes('QW|QW')) sheet.appendRow(HEADERS.QW_QW);
  }
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
