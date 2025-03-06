export interface OpenAIMessage {
  role: string;
  content: string;
}

export interface OpenAIChoice {
  text?: string;
  message?: {
    content: string;
  };
  delta?: {
    content: string;
  };
  index: number;
  finish_reason?: string;
}

export interface OpenAIResponse {
  id: string;
  object: string;
  created: number;
  model: string;
  choices: OpenAIChoice[];
}

export interface OpenAIRequestBody {
  model: string;
  messages?: OpenAIMessage[];
  prompt?: string;
  max_tokens: number;
  temperature?: number;
  stream?: boolean;
}

export interface APIError extends Error {
  status?: number;
  code?: string;
  response?: any;
}

export interface DataRange {
  sheetName: string;
  range: string;
  label?: string;
}

export interface SheetInfo {
  id: string;
  name: string;
  visible: boolean;
}

export interface RangeData {
  range: DataRange;
  values: any[][];
}

export interface MultiRangeData {
  ranges: RangeData[];
}

export interface ExcelService {
  getSheetsList(): Promise<SheetInfo[]>;
  getDataFromRange(range: string): Promise<any[][]>;
  getDataFromSheetRange(sheetName: string, range: string): Promise<any[][]>;
  getMultiRangeData(ranges: DataRange[]): Promise<MultiRangeData>;
}

export interface OpenAIService {
  analyze(prompt: string, onProgress?: (text: string) => void): Promise<string>;
}
