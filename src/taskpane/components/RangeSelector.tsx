import * as React from "react";
import {
  Stack,
  DefaultButton,
  IconButton,
  Label,
  CommandBarButton,
  IStackStyles,
  getTheme,
  IButtonStyles,
  Text
} from "@fluentui/react";
import { DataRange, SheetInfo } from "../../types/services";
import { ExcelService } from "../../services/excelService";

interface RangeSelectorProps {
  ranges: DataRange[];
  onRangesChange: (ranges: DataRange[]) => void;
}

const theme = getTheme();

const containerStyles: IStackStyles = {
  root: {
    padding: '8px',
    backgroundColor: theme.palette.neutralLighterAlt,
    borderRadius: '4px'
  }
};

const rangeItemStyles: IStackStyles = {
  root: {
    padding: '4px 8px',
    backgroundColor: theme.palette.white,
    borderRadius: '2px',
    border: `1px solid ${theme.palette.neutralLight}`,
    marginBottom: '4px'
  }
};

const buttonStyles: IButtonStyles = {
  root: {
    margin: '0 4px'
  }
};

export const RangeSelector: React.FC<RangeSelectorProps> = ({
  ranges,
  onRangesChange,
}) => {
  const [sheets, setSheets] = React.useState<SheetInfo[]>([]);
  const excelService = new ExcelService();

  React.useEffect(() => {
    loadSheets();
  }, []);

  const loadSheets = async () => {
    try {
      const sheetsList = await excelService.getSheetsList();
      setSheets(sheetsList.filter(s => s.visible));
    } catch (error) {
      console.error('Failed to load sheets:', error);
    }
  };

  const addRange = async () => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = context.workbook.getSelectedRange();
        sheet.load("name");
        range.load("address");
        await context.sync();

        const newRange: DataRange = {
          sheetName: sheet.name,
          range: range.address,
          label: `数据区域 ${ranges.length + 1}`
        };

        onRangesChange([...ranges, newRange]);
      });
    } catch (error) {
      console.error('Failed to add range:', error);
    }
  };

  const removeRange = (index: number) => {
    const newRanges = [...ranges];
    newRanges.splice(index, 1);
    onRangesChange(newRanges);
  };

  const updateRangeLabel = (index: number, newLabel: string) => {
    const newRanges = [...ranges];
    newRanges[index] = { ...newRanges[index], label: newLabel };
    onRangesChange(newRanges);
  };

  return (
    <Stack tokens={{ childrenGap: 8 }} styles={containerStyles}>
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
        <Label>数据区域</Label>
        <CommandBarButton
          iconProps={{ iconName: 'Add' }}
          text="添加数据区域"
          onClick={addRange}
          styles={buttonStyles}
        />
      </Stack>

      {ranges.map((range, index) => (
        <Stack key={index} styles={rangeItemStyles}>
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Stack tokens={{ childrenGap: 4 }}>
              <Text variant="mediumPlus">{range.label}</Text>
              <Text variant="small">
                工作表: {range.sheetName} | 区域: {range.range}
              </Text>
            </Stack>
            <IconButton
              iconProps={{ iconName: 'Delete' }}
              onClick={() => removeRange(index)}
              styles={buttonStyles}
            />
          </Stack>
        </Stack>
      ))}

      {ranges.length === 0 && (
        <Stack horizontalAlign="center" verticalAlign="center" styles={{ root: { padding: '16px' } }}>
          <Text>请选择一个数据区域并点击"添加数据区域"按钮</Text>
        </Stack>
      )}
    </Stack>
  );
};
