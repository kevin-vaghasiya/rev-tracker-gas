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

// interface IQWLead {
//   id: string;
//   timestamp: string;
//   sale_type: string;
//   property_address: string;
//   agent_name: string;
//   sale_price: number;
//   commission: number;
//   mgt_notes: string;
//   doc1_link: string;
//   mls_no: string;
//   settlement_date: string;
//   aos_date: string;
//   agent_2: string;
// }

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

interface IAgentsData {
  [key: string]: {
    q1: number;
    q2: number;
    q3: number;
    q4: number;
  };
}
