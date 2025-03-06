export class Validators {
  /**
   * Validates Excel range format
   * @param range Excel range string (e.g., "A1:B10")
   * @returns boolean indicating if the range format is valid
   */
  static isValidRange(range: string): boolean {
    // Basic range format validation (e.g., "A1:B2", "Sheet1!A1:B2")
    const rangeRegex = /^(?:[A-Za-z\d]+!)?[A-Z]+\d+(?::[A-Z]+\d+)?$/;
    return rangeRegex.test(range);
  }

  /**
   * Validates if the data is suitable for analysis
   * @param data The data array from Excel
   * @returns boolean indicating if the data is valid for analysis
   */
  static isValidAnalysisData(data: any[][]): boolean {
    if (!Array.isArray(data) || data.length === 0) {
      return false;
    }

    // Check if all rows have the same number of columns
    const columnCount = data[0].length;
    return data.every(row => 
      Array.isArray(row) && 
      row.length === columnCount && 
      row.every(cell => cell !== null && cell !== undefined)
    );
  }

  /**
   * Validates if the OpenAI API key is properly formatted
   * @param apiKey The OpenAI API key string
   * @returns boolean indicating if the API key format is valid
   */
  static isValidApiKey(apiKey: string): boolean {
    // Basic API key format validation
    return typeof apiKey === 'string' && apiKey.length > 0;
  }

  /**
   * Validates if the prompt is suitable for analysis
   * @param prompt The user's analysis prompt
   * @returns boolean indicating if the prompt is valid
   */
  static isValidPrompt(prompt: string): boolean {
    return typeof prompt === 'string' && 
           prompt.trim().length > 0 && 
           prompt.length <= 2000; // Reasonable length limit
  }
}
