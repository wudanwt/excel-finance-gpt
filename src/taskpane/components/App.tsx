import * as React from "react";
import { Stack, Text, getTheme, IStackStyles, DefaultButton, Label } from "@fluentui/react";
import { FinanceAnalyzer } from "./FinanceAnalyzer";
import Progress from "./Progress";

/* global console, Excel */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

const theme = getTheme();

const containerStyles: IStackStyles = {
  root: {
    height: '100vh',
    padding: 0,
    margin: 0,
    overflow: 'hidden',
    backgroundColor: theme.palette.white
  }
};

const innerContainerStyles: IStackStyles = {
  root: {
    height: '100%',
    overflow: 'hidden'
  }
};

const headerStyles: IStackStyles = {
  root: {
    padding: '8px 16px',
    borderBottom: `1px solid ${theme.palette.neutralLight}`
  }
};

const rangeSelectionStyles: IStackStyles = {
  root: {
    padding: '8px 16px',
    borderBottom: `1px solid ${theme.palette.neutralLight}`,
    backgroundColor: theme.palette.neutralLighterAlt
  }
};

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
    <Stack styles={containerStyles}>
      <Stack styles={innerContainerStyles}>
        <Stack.Item>
          <Stack styles={headerStyles}>
            <Text
              variant="xLarge"
              styles={{
                root: {
                  color: theme.palette.themePrimary,
                  fontWeight: 600
                }
              }}
            >
              Excel分析助手
            </Text>
          </Stack>
        </Stack.Item>

        <Stack.Item>
          <Stack horizontal verticalAlign="center" styles={rangeSelectionStyles} tokens={{ childrenGap: 16 }}>
            <DefaultButton
              onClick={handleGetCurrentRange}
              iconProps={{ iconName: 'TableComputed' }}
            >
              选择数据区域
            </DefaultButton>
            <Label
              styles={{
                root: {
                  color: theme.palette.themePrimary,
                  fontSize: '13px',
                  fontWeight: 400
                }
              }}
            >
              {selectedRange ? `当前选择：${selectedRange}` : '请选择要分析的数据区域'}
            </Label>
          </Stack>
        </Stack.Item>

        <Stack.Item grow>
          <FinanceAnalyzer
            range={selectedRange}
            isAnalyzing={isAnalyzing}
            setIsAnalyzing={setIsAnalyzing}
          />
        </Stack.Item>
      </Stack>
    </Stack>
  );
};
