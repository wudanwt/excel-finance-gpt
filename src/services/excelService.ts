import { ExcelService as IExcelService } from '../types/services';
import { logger } from '../utils/logger';
import { Validators } from '../utils/validators';

export class ExcelService implements IExcelService {
  async getDataFromRange(range: string): Promise<any[][]> {
    try {
      if (!Validators.isValidRange(range)) {
        throw new Error(`Invalid range format: ${range}`);
      }

      const result = await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const selectedRange = sheet.getRange(range);
        selectedRange.load("values");
        
        await context.sync();
        logger.debug('Excel data loaded from range:', range);
        return selectedRange.values;
      });

      if (!Validators.isValidAnalysisData(result)) {
        throw new Error('Invalid data format in selected range');
      }

      return result;
    } catch (err) {
      const error = err instanceof Error ? err : new Error('Unknown error reading Excel range');
      logger.error('Failed to get data from Excel range', error);
      throw error;
    }
  }

  private async validateSelectedRange(range: string): Promise<boolean> {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const selectedRange = sheet.getRange(range);
        selectedRange.load("rowCount,columnCount");
        await context.sync();
        
        if (selectedRange.rowCount === 0 || selectedRange.columnCount === 0) {
          throw new Error('Selected range is empty');
        }
      });
      return true;
    } catch {
      return false;
    }
  }
}
