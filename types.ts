interface IKeyData {
  [key: string]: {
    value: string;
    index: number;
    type: string;
    date?: string;
    email?: string;
  };
}

interface IHeaderIndexes {
  [key: string]: {
    index: number;
  };
}

interface IQWLead {
  Id: string;
  Timestamp: string;
  'Sale Type': string;
  'Property Address': string;
  'Agent Name': string;
  'Sale Price': number;
  'Commission Rate': number;
  MNotes: string;
  'Doc1 link': string;
  'MLS No': string;
  'Settlement Date': string;
  'AoS Date': string;
  QWAgent2: string;
}

interface ICache {
  date: {
    month: number;
    year: number;
  };
  sheet_name: string;
  type: string;
  index: number;
}
interface ICacheData {
  [key: string]: ICache;
}

interface ISheetCache {
  [key: string]: GoogleAppsScript.Spreadsheet.Sheet;
}

interface IHeadersCache {
  [key: string]: IHeaderIndexes;
}

interface IAgentsSplitData {
  [key: string]: {
    q1: number;
    q2: number;
    q3: number;
    q4: number;
  };
}

interface IAgentLead {
  Id: string;
  'Sale Type': string;
  'Property Address': string;
  'Sale Price': string;
  'Settlement Date': string;
  'AoS Date': string;
  Agent: string;
  Commission: number;
  Deductions: number;
  Revenue: number;
  QW_revenue: number;
}

interface IAgentsData {
  [key: string]: IAgentLead[];
}

interface ITotals {
  volume: number[];
  b_side: number[];
  s_side: number[];
  rev: number[];
  qw_rev: number[];
}
