import { config } from '../config/config';
import { logger } from '../utils/logger';
import { OpenAIService as IOpenAIService, OpenAIRequestBody, OpenAIResponse, APIError } from '../types/services';

export class OpenAIService implements IOpenAIService {
  private apiConfig = config.api.openai;

  private async makeRequest(endpoint: string, body: OpenAIRequestBody, onProgress?: (text: string) => void): Promise<OpenAIResponse> {
    const url = `${this.apiConfig.baseUrl}${endpoint}`;
    
    try {
      logger.debug('Making API request to:', url);
      logger.debug('Request body:', body);

      const response = await fetch(url, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${this.apiConfig.key}`,
          ...this.apiConfig.defaultHeaders
        },
        body: JSON.stringify({
          ...body,
          stream: !!onProgress
        })
      });

      if (!response.ok) {
        const errorData = await response.json();
        const error = new Error(errorData.error?.message || 'API request failed') as APIError;
        error.status = response.status;
        error.code = errorData.error?.code;
        error.response = errorData;
        throw error;
      }

      if (onProgress) {
        const reader = response.body!.getReader();
        const decoder = new TextDecoder();
        let buffer = '';
        let fullText = '';

        while (true) {
          const { done, value } = await reader.read();
          if (done) break;

          buffer += decoder.decode(value, { stream: true });
          const lines = buffer.split('\n');
          buffer = lines.pop() || '';

          for (const line of lines) {
            if (line.trim() === '') continue;
            if (line.trim() === 'data: [DONE]') break;
            if (!line.startsWith('data: ')) continue;

            try {
              const data = JSON.parse(line.slice(6));
              const text = data.choices?.[0]?.delta?.content || '';
              fullText += text;
              onProgress(fullText);
            } catch (e) {
              logger.error('Error parsing streaming response:', e);
            }
          }
        }

        return {
          id: 'stream',
          object: 'chat.completion',
          created: Date.now(),
          model: body.model,
          choices: [{
            message: { content: fullText },
            text: fullText,
            index: 0,
            finish_reason: 'stop'
          }]
        };
      }

      const responseData = await response.json();
      logger.debug('API response:', responseData);
      return responseData as OpenAIResponse;
    } catch (err) {
      logger.error('API request failed:', err);
      const error = err instanceof Error ? err : new Error('Unknown API error occurred');
      if (err instanceof Error) {
        logger.error('Error details:', {
          message: err.message,
          stack: err.stack,
          ...(err as any).response && { response: (err as any).response }
        });
      }
      throw error;
    }
  }

  public async analyze(prompt: string, onProgress?: (text: string) => void): Promise<string> {
    try {
      const requestBody: OpenAIRequestBody = {
        model: this.apiConfig.model,
        messages: [
          {
            role: "user",
            content: prompt
          }
        ],
        max_tokens: this.apiConfig.maxTokens,
        temperature: 0.7
      };

      const response = await this.makeRequest('/chat/completions', requestBody, onProgress);
      
      if (!response.choices || response.choices.length === 0) {
        throw new Error('No response from API');
      }

      const result = response.choices[0].message?.content || response.choices[0].text;
      if (!result) {
        throw new Error('Invalid response format from API');
      }

      logger.debug('Analysis result:', result);
      return result.trim();
    } catch (err) {
      const error = err instanceof Error ? err : new Error('Unknown error during analysis');
      logger.error('Analysis failed:', error);
      throw error;
    }
  }
}
