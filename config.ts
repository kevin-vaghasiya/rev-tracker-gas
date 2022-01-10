const SHEET_NAMES = {
  CONFIG: 'Config',
  LISTINGS: 'Listings',
  SPLITS: 'Splits',
  CACHE: 'Cache',
  ERROR_LOGS: 'error_logs',
};

const QWBT_SHEET_NAMES = {
  FORM_RESPONSES: 'Form Responses 1',
  NAME_MAPPING: 'Name Mapping',
  AGENT_EMAILS: 'Agent Emails',
  CHANGE_LOGS: 'Change_Logs',
};

const QW_KEY_NAMES = {
  TIMESTAMP: 'Timestamp',
  SALE_TYPE: 'Sale Type',
  PROPERTY_ADDRESS: 'Property Address',
  AGENT_NAME: 'Agent Name',
  SALE_PRICE: 'Sale Price',
  COMMISSION_TS: 'Commission Rate',
  MGT_NOTES: 'MNotes',
  DOC1: 'Doc1 link',
  MLS_NO: 'MLS No',
  ID: 'Id',
  SETTLEMENT_DATE: 'Settlement Date',

  // DRIVE_FOLDER_URL: 'Drive Folder Url',
  // SUBMITTED_TIME: 'Submitted Time',
  // SETTLEMENT_ADDRESS: 'Settlement Address',
  // DOC_LINK_0: 'Doc0 link',
  // DOC_LINK_1: 'Doc1 link',
  // DOC_LINK_2: 'Doc2 link',
  // DOC_LINK_3: 'Doc3 link',
  // DOC_LINK_4: 'Doc4 link',
  // DOC_LINK_5: 'Doc5 link',
  // DOC_LINK_6: 'Doc6 link',
  // DOC_LINK_7: 'Doc7 link',
  // FORM_EDIT_URL: 'Form Response Edit URL',
  // MORTGAGE_COMMIT_DATE: 'MCommit Date',
  // DEPOSIT_FIRST: '1 Dep Rec',
  // DEPOSIT_SECOND: '2 Dep Rec',
  // DEPOSIT_THIRD: '3 Dep Rec',
  // DEPOSIT_FIRST_DATE: '1 Dep Date',
  // DEPOSIT_SECOND_DATE: '2 Dep Date',
  // DEPOSIT_THIRD_DATE: '3 Dep Date',
  // EVENT_ID_1: 'Event Id 1',
  // EVENT_ID_2: 'Event Id 2',
  // EVENT_ID_3: 'Event Id 3',
  // EVENT_ID_4: 'Event Id 4',
  // EVENT_ID_5: 'Event Id 5',
  // CO_OP_AGENT_EMAIL: 'CO-OP Email',
  // CHANGED_FIELDS: 'Changed Fields',
  // CO_OP_UPDATE: 'UPDATE?',
};

const DEFAULT_SCRIPT_KEYS = {
  DRIVE_FOLDER_URL: 'Drive Folder Url',
  SUBMITTED_TIME: 'Submitted Time',
  DOC_LINK_0: 'Doc0 link',
  DOC_LINK_1: 'Doc1 link',
  DOC_LINK_2: 'Doc2 link',
  DOC_LINK_3: 'Doc3 link',
  DOC_LINK_4: 'Doc4 link',
  DOC_LINK_5: 'Doc5 link',
  DOC_LINK_6: 'Doc6 link',
  DOC_LINK_7: 'Doc7 link',
  FORM_EDIT_URL: 'Form Response Edit URL',
  TIMESTAMP: 'Timestamp',
  EVENT_ID_1: 'Event Id 1',
  EVENT_ID_2: 'Event Id 2',
  EVENT_ID_3: 'Event Id 3',
  EVENT_ID_4: 'Event Id 4',
  EVENT_ID_5: 'Event Id 5',
  ID: 'Id',
  CHANGED_FIELDS: 'Changed Fields',
};

const SALE_TYPES = {
  QW_LISTING_CO_OP: 'QW Listing | CO-OP Sale',
  QW_LISTING_QW: 'QW Listing | QW Sale',
  CO_OP_LISTING: 'CO-OP Listing | QW Sale',
  QW_LISTING_ACTIVE: 'QW Listing (Active/Coming Soon)',
  QW_NO_LONGER: 'QW No Longer Listed (WTH, EXP, etc)',
  FALL_THROUGH: 'Fallthrough Transaction',
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
