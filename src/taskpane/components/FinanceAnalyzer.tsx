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
  ITextFieldStyles
} from "@fluentui/react";
import { ExcelService } from "../../services/excelService";
import { OpenAIService } from "../../services/openaiService";
import { logger } from "../../utils/logger";

export interface FinanceAnalyzerProps {
  range: string;
  isAnalyzing: boolean;
  setIsAnalyzing: (analyzing: boolean) => void;
}

const stackTokens: IStackTokens = {
  childrenGap: 12,
  padding: '8px 4px'
};

const textFieldStyles: Partial<ITextFieldStyles> = {
  root: {
    width: '100%'
  },
  field: {
    height: '100%'
  }
};

const resultFieldStyles: Partial<ITextFieldStyles> = {
  ...textFieldStyles,
  root: {
    width: '100%',
    flex: 1,
    minHeight: '200px'
  },
  wrapper: {
    height: '100%'
  },
  field: {
    height: '100%',
    minHeight: '180px'
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
  const excelService = new ExcelService();
  const openaiService = new OpenAIService();

  const handleAnalyze = async () => {
    if (!range || !prompt) {
      setError("请选择数据区域并输入分析要求");
      return;
    }

    setIsAnalyzing(true);
    setError("");
    try {
      logger.debug('获取数据范围:', range);
      const data = await excelService.getDataFromRange(range);
      logger.debug('接收到数据:', data);

      const analysisPrompt = `${prompt}\n\n数据:\n${JSON.stringify(data)}`;
      logger.debug('发送到API的提示:', analysisPrompt);
      
      const analysis = await openaiService.analyze(analysisPrompt);
      logger.debug('收到分析结果:', analysis);
      
      setResult(analysis);
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
    <Stack tokens={stackTokens} styles={{ root: { height: '100%' } }}>
      {error && (
        <MessageBar messageBarType={MessageBarType.error}>
          {error}
        </MessageBar>
      )}
      
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
          disabled={isAnalyzing || !range || !prompt}
          iconProps={{ iconName: 'Analyze' }}
        >
          {isAnalyzing ? <Spinner size={SpinnerSize.small} /> : "开始分析"}
        </DefaultButton>
      </Stack.Item>

      <Stack.Item grow styles={{ root: { position: 'relative', minHeight: '200px' } }}>
        {result && (
          <TextField
            label="分析结果"
            multiline
            readOnly
            value={result}
            styles={resultFieldStyles}
          />
        )}
      </Stack.Item>
    </Stack>
  );
};
