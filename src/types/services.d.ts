export interface OpenAIMessage {
  role: string;
  content: string;
}

export interface OpenAIChoice {
  text?: string;
  message?: {
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
}

export interface APIError extends Error {
  status?: number;
  code?: string;
  response?: any;
}

export interface ExcelService {
  getDataFromRange(range: string): Promise<any[][]>;
}

export interface OpenAIService {
  analyze(prompt: string): Promise<string>;
}
