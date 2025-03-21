import * as React from "react";
import { createRoot } from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import ConfirmDialog from "./components/ConfirmDialog";

/* global document, Office, window */

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

Office.onReady(() => {
  // Get dialog data from sessionStorage
  let dialogData = { externalRecipients: [], attachments: [] };
  
  try {
    const storedData = window.sessionStorage.getItem('externalRecipientsData');
    if (storedData) {
      dialogData = JSON.parse(storedData);
    }
  } catch (error) {
    console.error("Error parsing dialog data:", error);
  }

  // Function to send message back to parent
  const sendMessageToParent = (message: any) => {
    Office.context.ui.messageParent(JSON.stringify(message));
  };

  root?.render(
    <FluentProvider theme={webLightTheme}>
      <ConfirmDialog 
        externalRecipients={dialogData.externalRecipients} 
        attachments={dialogData.attachments}
        onSend={() => sendMessageToParent({ action: 'send' })}
        onCancel={() => sendMessageToParent({ action: 'cancel' })}
      />
    </FluentProvider>
  );
});