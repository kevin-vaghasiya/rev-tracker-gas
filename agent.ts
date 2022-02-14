const updateAgentSheets = () => {
  const ss = SpreadsheetApp.getActive();
  const splitSheet = ss.getSheetByName(SHEET_NAMES.SPLITS);
  const split_data = splitSheet
    .getRange(1, 1, splitSheet.getLastRow(), splitSheet.getLastColumn())
    .getValues();
  const agents: IAgentsData = {};
  for (let i = 1; i < split_data.length; i++) {
    const [fun, fin, l_name] = split_data[i];
    if (!l_name) continue;
    agents[l_name] = [];
  }
  const TOTALS = {
    volume: new Array(12).fill(0),
    b_side: new Array(12).fill(0),
    s_side: new Array(12).fill(0),
    rev: new Array(12).fill(0),
    qw_rev: new Array(12).fill(0),
  };
  const leads: IAgentLead[] = getAllLeads(ss);
  for (let i = 0; i < leads.length; i++) {
    const { Agent } = leads[i];
    calculateDashboard(TOTALS, leads[i]);
    if (!Agent || !agents[Agent]) continue;
    agents[Agent].push(leads[i]);
  }
  setDashboard(ss, TOTALS);
  addAgentsData(ss, agents);
};

const setDashboard = (
  ss: GoogleAppsScript.Spreadsheet.Spreadsheet,
  TOTALS: ITotals
) => {
  const dashboard_sheet = ss.getSheetByName(SHEET_NAMES.DASHBOARD);
  const { volume, b_side, s_side, rev, qw_rev } = TOTALS;
  dashboard_sheet.getRange('B2:B13').setValues(volume.map((e) => [e]));
  dashboard_sheet.getRange('C2:C13').setValues(b_side.map((e) => [e]));
  dashboard_sheet.getRange('E2:E13').setValues(s_side.map((e) => [e]));
  dashboard_sheet.getRange('H2:H13').setValues(rev.map((e) => [e]));
  dashboard_sheet.getRange('I2:I13').setValues(qw_rev.map((e) => [e]));
};

const calculateDashboard = (TOTALS: ITotals, lead: IAgentLead) => {
  const { b_side, qw_rev, rev, s_side, volume } = TOTALS;
  const sale_type = lead['Sale Type'];
  const settlement_date = lead['Settlement Date'];
  const month = new Date(settlement_date).getMonth();
  const sale_price = Number(lead['Sale Price']);
  const Commission = lead['Commission'];
  const QW_revenue = Number(lead['QW_revenue']);

  if (sale_type == SALE_TYPES.CO_OP_LISTING) {
    volume[month] = volume[month] + sale_price;
    b_side[month] = b_side[month] + 1;
    rev[month] = rev[month] + Commission;
    qw_rev[month] = qw_rev[month] + QW_revenue;
  } else if (sale_type == SALE_TYPES.QW_LISTING_CO_OP) {
    volume[month] = volume[month] + sale_price;
    s_side[month] = s_side[month] + 1;
    rev[month] = rev[month] + Commission;
    qw_rev[month] = qw_rev[month] + QW_revenue;
  } else if (sale_type == SALE_TYPES.QW_LISTING_QW) {
    volume[month] = volume[month] + sale_price;
    b_side[month] = b_side[month] + 0.5;
    s_side[month] = s_side[month] + 0.5;
    rev[month] = rev[month] + Commission;
    qw_rev[month] = qw_rev[month] + QW_revenue;
  }
};

const addAgentsData = (
  ss: GoogleAppsScript.Spreadsheet.Spreadsheet,
  agents: IAgentsData
) => {
  for (const agent_name in agents) {
    if (!Object.prototype.hasOwnProperty.call(agents, agent_name)) continue;
    const agentLeads = agents[agent_name];

    let sheet = getAgentSheet(ss, agent_name);
    if (!agentLeads.length && !sheet) continue;
    if (sheet) {
      sheet.getRange('A2:Z').clear();
      if (!agentLeads.length) continue;
    } else sheet = createAgentSheet(ss, agent_name);

    const data = [];
    for (let i = 0; i < agentLeads.length; i++) {
      const {
        Id,
        'Sale Type': sale_type,
        'Property Address': property_address,
        'Settlement Date': settlement_date,
        'Sale Price': sale_price,
        'AoS Date': aos_date,
        Commission,
        Deductions,
        Revenue,
      } = agentLeads[i];
      data.push([
        sale_type,
        property_address,
        aos_date,
        settlement_date,
        sale_price,
        Commission,
        Deductions,
        Revenue,
        Id,
      ]);
    }
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  }
};

const getAllLeads = (ss: GoogleAppsScript.Spreadsheet.Spreadsheet) => {
  const sheets = getLeadSheets(ss);
  const leads: IAgentLead[] = [];

  // Gathering All Data For Data Studio Report
  const allData = [];
  const allSheet = ss.getSheetByName(SHEET_NAMES.ALL_DATA);
  allSheet.getRange('A2:Ae').clear();
  const all_header = getHeaderIndexes(allSheet);

  for (let i = 0; i < sheets.length; i++) {
    try {
      const sheet = sheets[i];
      const header = getHeaderIndexes(sheet);
      const data = sheet
        .getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn())
        .getValues();

      for (let j = 1; j < data.length; j++) {
        const [element] = data[j];
        if (!element) continue;
        const row = data[j];
        const Id = row[header[KEY_NAMES.ID]?.index];
        if (!Id) continue;

        mapAndAddAllData(allData, all_header, row, header);

        const sale_type = row[header[KEY_NAMES.SALE_TYPE]?.index];
        const { commission, deduction } = getCommissionAndDeductions(
          row,
          header
        );

        leads.push({
          Id,
          'Sale Type': sale_type,
          'Property Address': row[header[KEY_NAMES.PROPERTY_ADDRESS]?.index],
          'Sale Price': row[header[KEY_NAMES.SALE_PRICE].index],
          'Settlement Date': row[header[KEY_NAMES.SETTLEMENT_DATE]?.index],
          'AoS Date': row[header[KEY_NAMES.AOS_DATE]?.index],
          Commission: commission,
          Deductions: deduction,
          Revenue: row[header[KEY_NAMES.AGENT_REVENUE]?.index],
          Agent: row[header[KEY_NAMES.AGENT_NAME]?.index],
          QW_revenue: row[header[KEY_NAMES.QW_REVENUE]?.index],
          'List Price': row[header[KEY_NAMES.LIST_PRICE]?.index],
        });

        if (sale_type != SALE_TYPES.QW_LISTING_QW) continue;
        const { commission2, deduction2 } = getCommissionAndDeductions2(
          row,
          header
        );
        leads.push({
          Id,
          'Sale Type': sale_type,
          'Property Address': row[header[KEY_NAMES.PROPERTY_ADDRESS]?.index],
          'Sale Price': row[header[KEY_NAMES.SALE_PRICE].index],
          'Settlement Date': row[header[KEY_NAMES.SETTLEMENT_DATE]?.index],
          'AoS Date': row[header[KEY_NAMES.AOS_DATE]?.index],
          Commission: commission2,
          Deductions: deduction2,
          Revenue: row[header[KEY_NAMES.AGENT_REVENUE_2]?.index],
          Agent: row[header[KEY_NAMES.QW_AGENT_2]?.index],
          QW_revenue: row[header[KEY_NAMES.QW_REVENUE_2]?.index],
          'List Price': row[header[KEY_NAMES.LIST_PRICE]?.index],
        });
      }
    } catch (error) {
      console.log(error);
    }
  }
  if (allData.length)
    allSheet
      .getRange(2, 1, allData.length, allData[0].length)
      .setValues(allData);
  return leads;
};

const mapAndAddAllData = (
  allData: any[],
  all_header: IHeaderIndexes,
  row: string[],
  header: IHeaderIndexes
) => {
  const data = [];
  let flag = false;
  for (const key in header) {
    if (!Object.prototype.hasOwnProperty.call(header, key)) continue;
    const index = all_header[key].index;
    if (index || index == 0) {
      flag = true;
      data[index] = row[header[key].index];
    }
  }
  if (!flag) return;
  allData.push(data);
};

const getCommissionAndDeductions = (row: string[], header: IHeaderIndexes) => {
  const calculated_commission = Number(
    row[header[KEY_NAMES.CALCULATED_COMMISSION]?.index]
  );
  const manual_commission = Number(
    row[header[KEY_NAMES.MANUAL_COMMISSION]?.index] || 0
  );
  const commission =
    !isNaN(manual_commission) && manual_commission > 0
      ? manual_commission
      : calculated_commission;

  const automatic_deductions = Number(
    row[header[KEY_NAMES.AUTOMATIC_DEDUCTION]?.index]
  );
  const manual_deductions = Number(
    row[header[KEY_NAMES.MANUAL_DEDUCTION]?.index] || 0
  );
  const deduction =
    !isNaN(manual_deductions) && manual_deductions > 0
      ? manual_deductions
      : automatic_deductions;
  return { commission, deduction };
};

const getCommissionAndDeductions2 = (row: string[], header: IHeaderIndexes) => {
  let commission2 = 0;
  let deduction2 = 0;

  const calculated_commission = Number(
    row[header[KEY_NAMES.CALCULATED_COMMISSION_2]?.index]
  );
  const manual_commission = Number(
    row[header[KEY_NAMES.MANUAL_COMMISSION_2]?.index] || 0
  );
  commission2 =
    !isNaN(manual_commission) && manual_commission > 0
      ? manual_commission
      : calculated_commission;

  const automatic_deductions = Number(
    row[header[KEY_NAMES.AUTOMATIC_DEDUCTION_2]?.index]
  );
  const manual_deductions = Number(
    row[header[KEY_NAMES.MANUAL_DEDUCTION_2]?.index] || 0
  );
  deduction2 =
    !isNaN(manual_deductions) && manual_deductions > 0
      ? manual_deductions
      : automatic_deductions;
  return { commission2, deduction2 };
};

const getLeadSheets = (ss: GoogleAppsScript.Spreadsheet.Spreadsheet) => {
  const sheet_names = [];
  const sheets: GoogleAppsScript.Spreadsheet.Sheet[] = [];
  for (let i = 0; i < MONTH_NAMES.length; i++) {
    const month = MONTH_NAMES[i];
    sheet_names.push(`${month} QW|Co-Op`);
    sheet_names.push(`${month} Co-Op|QW`);
    sheet_names.push(`${month} QW|QW`);
  }
  for (let i = 0; i < sheet_names.length; i++) {
    const sheet_name = sheet_names[i];
    const sheet = ss.getSheetByName(sheet_name);
    if (!sheet) continue;
    sheets.push(sheet);
  }
  return sheets;
};

const getAgentSheet = (
  ss: GoogleAppsScript.Spreadsheet.Spreadsheet,
  agent_name: string
) => {
  return ss.getSheetByName(`Agent | ${agent_name}`);
};

const createAgentSheet = (
  ss: GoogleAppsScript.Spreadsheet.Spreadsheet,
  agent_name: string
) => {
  const sheet = ss.insertSheet(`Agent | ${agent_name}`);
  sheet.appendRow(HEADERS.AGENT);
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight('bold');
  return sheet;
};
