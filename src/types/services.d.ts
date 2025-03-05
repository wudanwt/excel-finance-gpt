// OpenAI API Types
export interface OpenAIRequestBody {
  model: string;
  prompt: string;
  max_tokens: number;
  temperature?: number;
  top_p?: number;
  frequency_penalty?: number;
  presence_penalty?: number;
}

export interface OpenAIChoice {
  text: string;
  index: number;
  logprobs: any;
  finish_reason: string;
}

export interface OpenAIResponse {
  id: string;
  object: string;
  created: number;
  model: string;
  choices: OpenAIChoice[];
  usage: {
    prompt_tokens: number;
    completion_tokens: number;
    total_tokens: number;
  };
}

// Service Interfaces
export interface OpenAIService {
  analyze(prompt: string): Promise<string>;
}

export interface ExcelService {
  getDataFromRange(range: string): Promise<any[][]>;
}

// Error Types
export interface APIError extends Error {
  status?: number;
  code?: string;
}

// Configuration Types
export interface APIConfig {
  baseUrl: string;
  key: string;
  timeout: number;
  maxTokens: number;
  model: string;
  defaultHeaders?: Record<string, string>;
}
