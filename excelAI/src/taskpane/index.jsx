import * as React from "react";
import { createRoot } from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import App from "./components/App";
/* global document, Office, module, require */

const title = "Excel AI Assistant";

// Get the container element
const container = document.getElementById("container");
const root = createRoot(container);

/* Ensures that the JS runs in the correct order */
Office.onReady(() => {
  root.render(
    <FluentProvider theme={webLightTheme}>
      <App title={title} />
    </FluentProvider>
  );
});

if (module.hot) {
  module.hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    root.render(
      <FluentProvider theme={webLightTheme}>
        <NextApp title={title} />
      </FluentProvider>
    );
  });
}
