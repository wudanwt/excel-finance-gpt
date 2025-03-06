import * as React from "react";
import { DefaultButton, Label, Stack, Text } from "@fluentui/react";
import { FinanceAnalyzer } from "./FinanceAnalyzer";
import Progress from "./Progress";

/* global console, Excel */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export const App: React.FC<AppProps> = ({ title, isOfficeInitialized }) => {
  const [isAnalyzing, setIsAnalyzing] = React.useState(false);
  const [selectedRange, setSelectedRange] = React.useState("");

  const handleGetCurrentRange = async () => {
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load("address");
        await context.sync();
        setSelectedRange(range.address);
      });
    } catch (error) {
      console.error("获取选定范围时出错:", error);
    }
  };

  if (!isOfficeInitialized) {
    return (
      <Progress
        title={title}
        logo="assets/logo-filled.png"
        message="正在初始化加载项，请稍候..."
      />
    );
  }

  return (
    <Stack tokens={{ padding: '8px', childrenGap: 8 }}>
      <Stack.Item>
        <Text variant="xLarge" block styles={{
          root: {
            fontWeight: '600',
            marginBottom: '8px'
          }
        }}>
          Excel 财务分析助手
        </Text>
      </Stack.Item>

      <Stack.Item>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
          <DefaultButton 
            onClick={handleGetCurrentRange}
            iconProps={{ iconName: 'Table' }}
          >
            选取数据区域
          </DefaultButton>
          {selectedRange && (
            <Label styles={{
              root: {
                marginLeft: '8px',
                fontSize: '14px',
                color: '#217346'
              }
            }}>
              当前选择: {selectedRange}
            </Label>
          )}
        </Stack>
      </Stack.Item>

      <Stack.Item grow>
        {selectedRange && (
          <FinanceAnalyzer 
            range={selectedRange}
            isAnalyzing={isAnalyzing}
            setIsAnalyzing={setIsAnalyzing}
          />
        )}
      </Stack.Item>
    </Stack>
  );
};
