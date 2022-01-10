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
  id: string;
  timestamp: string;
  sale_type: string;
  property_address: string;
  agent_name: string;
  sale_price: number;
  commission: number;
  mgt_notes: string;
  doc1_link: string;
  mls_no: string;
  settlement_date: string;
}

interface ICache {
  [key: string]: {
    date: {
      month: number;
      year: number;
    };
    sheet_name: string;
    type: string;
    index: number;
  };
}

interface ISheetCache {
  [key: string]: GoogleAppsScript.Spreadsheet.Sheet;
}
