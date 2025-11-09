export interface ExcelFile {
  id: string;
  name: string;
  webUrl: string;
  createdDateTime: string;
  lastModifiedDateTime: string;
  size: number;
}

export interface Worksheet {
  id: string;
  name: string;
  position: number;
  visibility: string;
}

export interface ExcelData {
  values: any[][];
  address: string;
  rowCount: number;
  columnCount: number;
}

export interface RowData {
  [key: string]: any;
}

export interface CreateRowPayload {
  driveId: string;
  itemId: string;
  worksheetName: string;
  data: RowData;
}

export interface UpdateRowPayload extends CreateRowPayload {
  rowIndex: string;
}

export interface DeleteRowPayload {
  driveId: string;
  itemId: string;
  worksheetName: string;
  rowIndex: string;
}

export interface ExcelContext {
  driveId: string | null;
  itemId: string | null;
  worksheetName: string | null;
}
