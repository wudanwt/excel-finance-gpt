import { logger } from './logger';

export class Validators {
  /**
   * Validates an Excel range string format (e.g., "A1:B10")
   */
  static isValidRange(range: string): boolean {
    const rangePattern = /^[A-Za-z]+[0-9]+:[A-Za-z]+[0-9]+$/;
    const isValid = rangePattern.test(range);
    if (!isValid) {
      logger.warn(`Invalid range format: ${range}`);
    }
    return isValid;
  }

  /**
   * Validates if the value is a non-empty string
   */
  static isNonEmptyString(value: any): boolean {
    return typeof value === 'string' && value.trim().length > 0;
  }

  /**
   * Validates an API key format
   */
  static isValidApiKey(apiKey: string): boolean {
    // OpenAI API keys typically start with 'sk-' and are 51 characters long
    const isValid = apiKey.startsWith('sk-') && apiKey.length === 51;
    if (!isValid) {
      logger.warn('Invalid API key format');
    }
    return isValid;
  }

  /**
   * Validates the analysis data format
   */
  static isValidAnalysisData(data: any[][]): boolean {
    if (!Array.isArray(data) || data.length === 0) {
      logger.warn('Invalid data format: empty or not an array');
      return false;
    }

    // Check if all rows have the same number of columns
    const columnCount = data[0].length;
    const isValid = data.every(row => Array.isArray(row) && row.length === columnCount);
    
    if (!isValid) {
      logger.warn('Invalid data format: inconsistent row lengths');
    }
    
    return isValid;
  }

  /**
   * Validates the analysis type
   */
  static isValidAnalysisType(type: string): boolean {
    const validTypes = ['financial', 'trend', 'risk'];
    const isValid = validTypes.includes(type.toLowerCase());
    if (!isValid) {
      logger.warn(`Invalid analysis type: ${type}`);
    }
    return isValid;
  }
}
