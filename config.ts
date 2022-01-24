const SHEET_NAMES = {
  CONFIG: 'Config',
  LISTINGS: 'Listings',
  SPLITS: 'Splits',
  CACHE: 'Cache',
  ERROR_LOGS: 'error_logs',
  ALL_DATA: 'All Data',
  LOGS: 'Logs',
};

const QWBT_SHEET_NAMES = {
  FORM_RESPONSES: 'Form Responses 1',
  NAME_MAPPING: 'Name Mapping',
  AGENT_EMAILS: 'Agent Emails',
  CHANGE_LOGS: 'Change_Logs',
};

const KEY_NAMES = {
  TIMESTAMP: 'Timestamp',
  SALE_TYPE: 'Sale Type',
  PROPERTY_ADDRESS: 'Property Address',
  AGENT_NAME: 'Agent Name',
  SALE_PRICE: 'Sale Price',
  COMMISSION_RATE: 'Commission Rate',
  MGT_NOTES: 'MGT Notes Link',
  DOC1: 'Doc 1 Link',
  MLS_NO: 'MLS No',
  ID: 'Id',
  SETTLEMENT_DATE: 'Settlement Date',
  AOS_DATE: 'AoS Date',
  QW_AGENT_2: 'QW Agent 2',
  PROJECTED_REV: 'Projected Rev',
  CALCULATED_COMMISSION: 'Calculated Commission',
  MANUAL_COMMISSION: 'Manual Commission',
  AUTOMATIC_DEDUCTION: 'Automatic Deduction',
  MANUAL_DEDUCTION: 'Manual Deductions',
  REVENUE_FOR_SPLIT: 'Revenue for Split',
  AGENT_REVENUE: 'Agent Revenue',
  QW_REVENUE: 'QW Revenue',
  COMMISSION_2: 'Commission Rate 2',
  CALCULATED_COMMISSION_2: 'Calculated Commission 2',
  MANUAL_COMMISSION_2: 'Manual Commission 2',
  AUTOMATIC_DEDUCTION_2: 'Automatic Deduction 2',
  MANUAL_DEDUCTION_2: 'Manual Deductions 2',
  REVENUE_FOR_SPLIT_2: 'Revenue for Split 2',
  AGENT_REVENUE_2: 'Agent Revenue 2',
  QW_REVENUE_2: 'QW Revenue 2',
};

const MANUAL_FIELDS = [
  KEY_NAMES.MANUAL_COMMISSION,
  KEY_NAMES.MANUAL_DEDUCTION,
  KEY_NAMES.MANUAL_COMMISSION_2,
  KEY_NAMES.MANUAL_DEDUCTION_2,
];

const QW_KEY_NAMES = {
  TIMESTAMP: 'Timestamp',
  SALE_TYPE: 'Sale Type',
  PROPERTY_ADDRESS: 'Property Address',
  AGENT_NAME: 'Agent Name',
  SALE_PRICE: 'Sale Price',
  COMMISSION_RATE: 'Commission Rate',
  MGT_NOTES: 'MNotes',
  DOC1: 'Doc1 link',
  MLS_NO: 'MLS No',
  ID: 'Id',
  SETTLEMENT_DATE: 'Settlement Date',
  AOS_DATE: 'AoS Date',
  QW_AGENT_2: 'QWAgent2',
  COMMISSION_2: 'Commission 2',
};

const SALE_TYPES = {
  QW_LISTING_CO_OP: 'QW Listing | CO-OP Sale',
  QW_LISTING_QW: 'QW Listing | QW Sale',
  CO_OP_LISTING: 'CO-OP Listing | QW Sale',
  QW_LISTING_ACTIVE: 'QW Listing (Active/Coming Soon)',
  QW_NO_LONGER: 'QW No Longer Listed (WTH, EXP, etc)',
  FALL_THROUGH: 'Fallthrough Transaction',
};

const AUTOMATIC_DEDUCTIONS = {
  CO_OP_QW: 72.5,
  QW_CO_OP: 60,
  QW_QW: 30,
};

const SEND_EMAIL_AFTER_HOURS = 24;

const MONTH_NAMES = [
  'Jan',
  'Feb',
  'Mar',
  'Apr',
  'May',
  'Jun',
  'Jul',
  'Aug',
  'Sep',
  'Oct',
  'Nov',
  'Dec',
];

const HEADERS = {
  LISTING: [
    'Timestamp',
    'Sale Type',
    'Property Address',
    'Agent Name',
    'Sale Price',
    'Commission Rate',
    'Projected Rev',
    'MGT Notes Link',
    'Doc 1 Link',
    'MLS No',
    'Id',
  ],
  QW_CO_OP: [
    'Timestamp',
    'Settlement Date',
    'Sale Type',
    'Property Address',
    'Agent Name',
    'Sale Price',
    'AoS Date',
    'Commission Rate',
    'Calculated Commission',
    'Manual Commission',
    'Automatic Deduction',
    'Manual Deductions',
    'Revenue for Split',
    'Agent Revenue',
    'QW Revenue',
    'MGT Notes Link',
    'Doc 1 Link',
    'Id',
  ],
  CO_OP_QW: [
    'Timestamp',
    'Settlement Date',
    'Sale Type',
    'Property Address',
    'Agent Name',
    'Sale Price',
    'AoS Date',
    'Commission Rate',
    'Calculated Commission',
    'Manual Commission',
    'Automatic Deduction',
    'Manual Deductions',
    'Revenue for Split',
    'Agent Revenue',
    'QW Revenue',
    'MGT Notes Link',
    'Doc 1 Link',
    'MLS No',
    'Id',
  ],
  QW_QW: [
    'Timestamp',
    'Settlement Date',
    'Sale Type',
    'Property Address',
    'Agent Name',
    'Sale Price',
    'AoS Date',
    'Commission Rate',
    'Calculated Commission',
    'Manual Commission',
    'Automatic Deduction',
    'Manual Deductions',
    'Revenue for Split',
    'Agent Revenue',
    'QW Revenue',
    'QW Agent 2',
    'Commission Rate 2',
    'Calculated Commission 2',
    'Manual Commission 2',
    'Automatic Deduction 2',
    'Manual Deductions 2',
    'Revenue for Split 2',
    'Agent Revenue 2',
    'QW Revenue 2',
    'MGT Notes Link',
    'Doc 1 Link',
    'MLS No',
    'Id',
  ],
  AGENT: [
    'Sale Type',
    'Property Address',
    'AoS Date',
    'Settlement Date',
    'Sale Price',
    'Commission',
    'Deductions',
    'Revenue',
    'Id',
  ],
};
