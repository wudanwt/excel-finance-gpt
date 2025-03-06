import * as React from "react";
import * as ReactDOM from "react-dom";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import { App } from "./components/App";
/* global document, Office, module, require */

initializeIcons();

let isOfficeInitialized = false;

const title = "Excel Finance GPT";

const render = (Component: typeof App) => {
  ReactDOM.render(
    <Component title={title} isOfficeInitialized={isOfficeInitialized} />,
    document.getElementById("container")
  );
};

/* Render application after Office initializes */
Office.onReady(() => {
  isOfficeInitialized = true;
  render(App);
}).catch(error => {
  console.error('Error while initializing Office:', error);
});

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").App;
    render(NextApp);
  });
}
