import * as React from 'react';
import { Stack, PrimaryButton, TextField, ComboBox, IComboBox, IComboBoxOption, ProgressIndicator, MessageBar, MessageBarType } from '@fluentui/react';
import { ExcelService, OpenAIService } from '../../types/services';

export type AnalysisType = 'financial' | 'trend' | 'risk';

interface FinanceAnalyzerProps {
  excelService: ExcelService;
  openaiService: OpenAIService;
}

const analysisOptions: IComboBoxOption[] = [
  { key: 'financial', text: 'Financial Analysis' },
  { key: 'trend', text: 'Trend Analysis' },
  { key: 'risk', text: 'Risk Analysis' }
];

export const FinanceAnalyzer: React.FC<FinanceAnalyzerProps> = ({ excelService, openaiService }) => {
  const [selectedRange, setSelectedRange] = React.useState('');
  const [analysisType, setAnalysisType] = React.useState<AnalysisType>('financial');
  const [isAnalyzing, setIsAnalyzing] = React.useState(false);
  const [result, setResult] = React.useState('');
  const [error, setError] = React.useState('');

  const handleRangeChange = (_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
    if (newValue !== undefined) {
      setSelectedRange(newValue);
      setError('');
    }
  };

  const handleAnalysisTypeChange = (_: React.FormEvent<IComboBox>, option?: IComboBoxOption) => {
    if (option) {
      setAnalysisType(option.key as AnalysisType);
      setError('');
    }
  };

  const handleAnalyze = async () => {
    if (!selectedRange) {
      setError('Please select a range');
      return;
    }

    setIsAnalyzing(true);
    setError('');
    setResult('');
    
    try {
      const data = await excelService.getDataFromRange(selectedRange);
      const prompt = `Analyze the following financial data using ${analysisType} analysis method:\n${JSON.stringify(data)}`;
      const analysis = await openaiService.analyze(prompt);
      setResult(analysis);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'An error occurred during analysis');
    } finally {
      setIsAnalyzing(false);
    }
  };

  return (
    <Stack tokens={{ childrenGap: 15 }}>
      <TextField
        label="Data Range"
        value={selectedRange}
        onChange={handleRangeChange}
        placeholder="e.g., A1:B10"
        disabled={isAnalyzing}
      />
      
      <ComboBox
        label="Analysis Type"
        selectedKey={analysisType}
        options={analysisOptions}
        onChange={handleAnalysisTypeChange}
        disabled={isAnalyzing}
      />

      <PrimaryButton
        text="Analyze"
        onClick={handleAnalyze}
        disabled={isAnalyzing || !selectedRange}
      />

      {isAnalyzing && (
        <ProgressIndicator label="Analyzing data..." />
      )}

      {error && (
        <MessageBar messageBarType={MessageBarType.error}>
          {error}
        </MessageBar>
      )}

      {result && !error && !isAnalyzing && (
        <Stack tokens={{ childrenGap: 10 }}>
          <MessageBar messageBarType={MessageBarType.success}>
            Analysis completed
          </MessageBar>
          <div style={{ whiteSpace: 'pre-wrap' }}>
            {result}
          </div>
        </Stack>
      )}
    </Stack>
  );
};
