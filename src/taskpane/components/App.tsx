import * as React from 'react';
import { Stack } from '@fluentui/react';
import Progress from './Progress';
import { FinanceAnalyzer } from './FinanceAnalyzer';
import { ExcelService as ExcelServiceImpl } from '../../services/excelService';
import { OpenAIService as OpenAIServiceImpl } from '../../services/openaiService';
import { ExcelService, OpenAIService } from '../../types/services';

/* global require */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

const excelService: ExcelService = new ExcelServiceImpl();
const openaiService: OpenAIService = new OpenAIServiceImpl();

export default function App({ title, isOfficeInitialized }: AppProps) {
  if (!isOfficeInitialized) {
    return (
      <Progress
        title={title}
        logo={require('./../../../assets/logo-filled.png')}
        message="Please sideload your addin to see app body."
      />
    );
  }

  return (
    <Stack className="ms-welcome" tokens={{ childrenGap: 10 }}>
      <h2>{title}</h2>
      <FinanceAnalyzer
        excelService={excelService}
        openaiService={openaiService}
      />
    </Stack>
  );
}
