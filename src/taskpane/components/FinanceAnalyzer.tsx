import * as React from "react";
import {
  DefaultButton,
  Spinner,
  SpinnerSize,
  Stack,
  TextField,
  MessageBar,
  MessageBarType,
  IStackTokens,
  ITextFieldStyles,
  getTheme,
  IStackStyles
} from "@fluentui/react";
import { ExcelService } from "../../services/excelService";
import { OpenAIService } from "../../services/openaiService";
import { logger } from "../../utils/logger";
import { RangeSelector } from "./RangeSelector";
import { DataRange } from "../../types/services";

export interface FinanceAnalyzerProps {
  range?: string;
  isAnalyzing: boolean;
  setIsAnalyzing: (analyzing: boolean) => void;
}

const theme = getTheme();

const stackTokens: IStackTokens = {
  childrenGap: 12,
  padding: '8px'
};

const analyzerContainerStyles: IStackStyles = {
  root: {
    height: '100%',
    display: 'flex',
    flexDirection: 'column'
  }
};

const textFieldStyles: Partial<ITextFieldStyles> = {
  root: {
    width: '100%'
  },
  field: {
    height: '100%',
    minHeight: '60px'
  }
};

const resultContainerStyles: IStackStyles = {
  root: {
    flex: 1,
    display: 'flex',
    flexDirection: 'column',
    minHeight: 0,
    marginBottom: '8px'
  }
};

const resultFieldStyles: Partial<ITextFieldStyles> = {
  root: {
    height: '100%',
    flex: 1,
    display: 'flex',
    flexDirection: 'column'
  },
  wrapper: {
    height: '100%',
    display: 'flex',
    flexDirection: 'column'
  },
  fieldGroup: {
    flex: 1,
    height: 'auto'
  },
  field: {
    height: '100%',
    padding: '8px',
    resize: 'none',
    backgroundColor: theme.palette.white,
    border: `1px solid ${theme.palette.neutralLight}`,
    borderRadius: '2px',
    overflowY: 'auto'
  }
};

export const FinanceAnalyzer: React.FC<FinanceAnalyzerProps> = ({
  range,
  isAnalyzing,
  setIsAnalyzing,
}) => {
  const [result, setResult] = React.useState<string>("");
  const [prompt, setPrompt] = React.useState<string>("");
  const [error, setError] = React.useState<string>("");
  const [ranges, setRanges] = React.useState<DataRange[]>([]);
  const excelService = new ExcelService();
  const openaiService = new OpenAIService();
  const resultRef = React.useRef<HTMLDivElement>(null);

  // 处理传统单区域模式
  React.useEffect(() => {
    if (range && ranges.length === 0) {
      handleLegacyRange();
    }
  }, [range]);

  const handleLegacyRange = async () => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.load("name");
        await context.sync();

        setRanges([{
          sheetName: sheet.name,
          range: range!,
          label: "数据区域 1"
        }]);
      });
    } catch (error) {
      console.error('Failed to handle legacy range:', error);
    }
  };

  // 自动滚动到底部
  React.useEffect(() => {
    if (resultRef.current) {
      resultRef.current.scrollTop = resultRef.current.scrollHeight;
    }
  }, [result]);

  const handleAnalyze = async () => {
    if (ranges.length === 0) {
      setError("请至少选择一个数据区域");
      return;
    }

    if (!prompt) {
      setError("请输入分析要求");
      return;
    }

    setIsAnalyzing(true);
    setError("");
    setResult("");

    try {
      logger.debug('正在获取多个数据区域的数据');
      const multiRangeData = await excelService.getMultiRangeData(ranges);
      logger.debug('接收到数据:', multiRangeData);

      const formattedData = multiRangeData.ranges.map(rangeData => ({
        label: rangeData.range.label || `${rangeData.range.sheetName}!${rangeData.range.range}`,
        sheetName: rangeData.range.sheetName,
        range: rangeData.range.range,
        data: rangeData.values
      }));

      const analysisPrompt = `
分析要求: ${prompt}

数据区域:
${formattedData.map(data => `
${data.label} (${data.sheetName}!${data.range}):
${JSON.stringify(data.data, null, 2)}`).join('\n')}
`;

      logger.debug('发送到API的提示:', analysisPrompt);
      
      await openaiService.analyze(
        analysisPrompt,
        (text) => {
          setResult(text);
        }
      );
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : '发生未知错误';
      logger.error('分析失败:', error);
      setError(`分析失败: ${errorMessage}`);
      setResult("分析失败，请查看控制台了解详细信息。");
    } finally {
      setIsAnalyzing(false);
    }
  };

  return (
    <Stack styles={analyzerContainerStyles} tokens={stackTokens}>
      {error && (
        <Stack.Item>
          <MessageBar messageBarType={MessageBarType.error}>
            {error}
          </MessageBar>
        </Stack.Item>
      )}
      
      <Stack.Item>
        <RangeSelector ranges={ranges} onRangesChange={setRanges} />
      </Stack.Item>

      <Stack.Item>
        <TextField
          label="分析要求"
          multiline
          rows={3}
          value={prompt}
          onChange={(_, newValue) => setPrompt(newValue || "")}
          placeholder="请输入您要分析的内容..."
          styles={textFieldStyles}
        />
      </Stack.Item>

      <Stack.Item>
        <DefaultButton
          onClick={handleAnalyze}
          disabled={isAnalyzing || ranges.length === 0 || !prompt}
          iconProps={{ iconName: 'AnalyticsReport' }}
        >
          {isAnalyzing ? <Spinner size={SpinnerSize.small} /> : "开始分析"}
        </DefaultButton>
      </Stack.Item>

      <Stack.Item grow styles={resultContainerStyles}>
        <TextField
          label="分析结果"
          multiline
          readOnly
          value={result}
          styles={resultFieldStyles}
          componentRef={resultRef as any}
        />
      </Stack.Item>
    </Stack>
  );
};
