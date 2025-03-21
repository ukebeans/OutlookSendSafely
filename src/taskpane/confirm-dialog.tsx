import * as React from "react";
import { createRoot } from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import ConfirmDialog from "./components/ConfirmDialog";

/* global document, Office, window */

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

Office.onReady(() => {
  // Get dialog data from sessionStorage or directly from URL parameters
  let dialogData = { externalRecipients: [], attachments: [] };
  
  try {
    // First try to get data from sessionStorage
    const storedData = window.sessionStorage.getItem('externalRecipientsData');
    if (storedData) {
      dialogData = JSON.parse(storedData);
      console.log("Retrieved dialog data from sessionStorage");
    } else {
      // If no data in sessionStorage, check for URL parameters (fallback for OWA)
      console.log("No data found in sessionStorage, checking URL parameters");
      const urlParams = new URLSearchParams(window.location.search);
      const dataParam = urlParams.get('data');
      if (dataParam) {
        try {
          // URL params are usually encoded
          const decodedData = decodeURIComponent(dataParam);
          dialogData = JSON.parse(decodedData);
          console.log("Retrieved dialog data from URL parameters");
        } catch (e) {
          console.error("Error parsing URL parameter data:", e);
        }
      }
    }
  } catch (error) {
    console.error("Error retrieving dialog data:", error);
  }

  // Log whether we have data or not
  if (dialogData.externalRecipients.length > 0) {
    console.log(`Dialog data loaded with ${dialogData.externalRecipients.length} external recipients`);
  } else {
    console.warn("No external recipients found in dialog data");
  }

  // Function to send message back to parent
  const sendMessageToParent = (message: any) => {
    try {
      Office.context.ui.messageParent(JSON.stringify(message));
      console.log("Message sent to parent:", message);
    } catch (error) {
      console.error("Error sending message to parent:", error);
    }
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