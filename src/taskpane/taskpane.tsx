import App from "./components/App";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import { ThemeProvider } from "@fluentui/react";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { logger } from "../utils/logger";

/* global document, Office, module, require */

initializeIcons();

let isOfficeInitialized = false;

const title = "Excel Finance GPT";

interface ComponentType {
  (props: { title: string; isOfficeInitialized: boolean }): JSX.Element;
}

const render = (Component: ComponentType): void => {
  ReactDOM.render(
    <AppContainer>
      <ThemeProvider>
        <Component title={title} isOfficeInitialized={isOfficeInitialized} />
      </ThemeProvider>
    </AppContainer>,
    document.getElementById("container")
  );
};

/* Render application after Office initializes */
Office.onReady(() => {
  isOfficeInitialized = true;
  logger.info("Office.js is ready");
  render(App);
}).catch((error) => {
  logger.error("Error during Office.js initialization:", error);
});

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(NextApp);
  });
}
