import { config } from '../config/config';
import { logger } from '../utils/logger';
import { OpenAIService as IOpenAIService, OpenAIRequestBody, OpenAIResponse, APIError } from '../types/services';

export class OpenAIService implements IOpenAIService {
  private apiConfig = config.api.openai;

  private async makeRequest(endpoint: string, body: OpenAIRequestBody): Promise<OpenAIResponse> {
    const url = `${this.apiConfig.baseUrl}${endpoint}`;
    
    try {
      const response = await fetch(url, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${this.apiConfig.key}`,
          ...this.apiConfig.defaultHeaders
        },
        body: JSON.stringify(body)
      });

      if (!response.ok) {
        const errorData = await response.json();
        const error = new Error(errorData.error?.message || 'API request failed') as APIError;
        error.status = response.status;
        error.code = errorData.error?.code;
        throw error;
      }

      const result = await response.json();
      return result as OpenAIResponse;
    } catch (err) {
      const error = err instanceof Error ? err : new Error('Unknown API error occurred');
      logger.error('API request failed', error);
      throw error;
    }
  }

  public async analyze(prompt: string): Promise<string> {
    try {
      const requestBody: OpenAIRequestBody = {
        model: this.apiConfig.model,
        prompt,
        max_tokens: this.apiConfig.maxTokens,
        temperature: 0.7
      };

      const response = await this.makeRequest('/completions', requestBody);
      
      if (!response.choices || response.choices.length === 0) {
        throw new Error('No response from API');
      }

      logger.debug('API response received', response);
      return response.choices[0].text.trim();
    } catch (err) {
      const error = err instanceof Error ? err : new Error('Unknown error during analysis');
      logger.error('Analysis failed', error);
      throw error;
    }
  }
}
