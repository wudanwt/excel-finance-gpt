import { ExcelService as IExcelService, SheetInfo, DataRange, MultiRangeData, RangeData } from '../types/services';
import { logger } from '../utils/logger';
import { Validators } from '../utils/validators';

export class ExcelService implements IExcelService {
  async getSheetsList(): Promise<SheetInfo[]> {
    try {
      const result = await Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        sheets.load(['items/name', 'items/id', 'items/visibility']);
        await context.sync();
        
        return sheets.items.map(sheet => ({
          id: sheet.id,
          name: sheet.name,
          visible: sheet.visibility === Excel.SheetVisibility.visible
        }));
      });

      return result;
    } catch (err) {
      const error = err instanceof Error ? err : new Error('Unknown error getting sheets list');
      logger.error('Failed to get sheets list', error);
      throw error;
    }
  }

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

  async getDataFromSheetRange(sheetName: string, range: string): Promise<any[][]> {
    try {
      if (!Validators.isValidRange(range)) {
        throw new Error(`Invalid range format: ${range}`);
      }

      const result = await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        const selectedRange = sheet.getRange(range);
        selectedRange.load("values");
        
        await context.sync();
        logger.debug(`Excel data loaded from sheet: ${sheetName}, range: ${range}`);
        return selectedRange.values;
      });

      if (!Validators.isValidAnalysisData(result)) {
        throw new Error('Invalid data format in selected range');
      }

      return result;
    } catch (err) {
      const error = err instanceof Error ? err : new Error(`Unknown error reading Excel range from sheet: ${sheetName}`);
      logger.error('Failed to get data from Excel sheet range', error);
      throw error;
    }
  }

  async getMultiRangeData(ranges: DataRange[]): Promise<MultiRangeData> {
    try {
      const result: MultiRangeData = {
        ranges: []
      };

      const rangesData = await Excel.run(async (context) => {
        const rangesPromises = ranges.map(dataRange => {
          const sheet = context.workbook.worksheets.getItem(dataRange.sheetName);
          const range = sheet.getRange(dataRange.range);
          range.load("values");
          return range;
        });

        // 等待所有范围加载完成
        await context.sync();

        // 收集所有数据
        return rangesPromises.map((range, index) => {
          if (!Validators.isValidAnalysisData(range.values)) {
            throw new Error(`Invalid data format in range: ${ranges[index].range} from sheet: ${ranges[index].sheetName}`);
          }

          return {
            range: ranges[index],
            values: range.values
          };
        });
      });

      result.ranges = rangesData;
      return result;

    } catch (err) {
      const error = err instanceof Error ? err : new Error('Unknown error reading multiple Excel ranges');
      logger.error('Failed to get data from multiple Excel ranges', error);
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
