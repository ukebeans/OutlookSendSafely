/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event: Office.AddinCommands.Event) {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message.
  Office.context.mailbox.item?.notificationMessages.replaceAsync(
    "ActionPerformanceNotification",
    message
  );

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

/**
 * Validates if email is being sent to external recipients and displays a confirmation dialog
 * @param event The event object from Outlook
 */
function validateExternalRecipients(event: Office.AddinCommands.Event) {
  const item = Office.context.mailbox.item;
  
  // Function to check if an email domain is external
  const isExternalDomain = (email: string): boolean => {
    // Check if the email ends with your organization's domain
    return !email.toLowerCase().endsWith("@outlook.com");
  };

  // Get all recipients - using a more generic approach
  const recipients: Office.EmailAddressDetails[] = [];
  
  // Check To recipients
  item.to?.getAsync((toResult) => {
    if (toResult.status === Office.AsyncResultStatus.Succeeded) {
      const toRecipients = toResult.value || [];
      recipients.push(...toRecipients);
    }
    
    // Check CC recipients
    item.cc?.getAsync((ccResult) => {
      if (ccResult.status === Office.AsyncResultStatus.Succeeded) {
        const ccRecipients = ccResult.value || [];
        recipients.push(...ccRecipients);
      }
      
      // Check BCC recipients
      item.bcc?.getAsync((bccResult) => {
        if (bccResult.status === Office.AsyncResultStatus.Succeeded) {
          const bccRecipients = bccResult.value || [];
          recipients.push(...bccRecipients);
        }
        
        // Process all gathered recipients
        const externalRecipients = recipients.filter(recipient => {
          try {
            return isExternalDomain(recipient.emailAddress);
          } catch (e) {
            console.error("Error checking email domain", e);
            return false;
          }
        });
        
        // If there are no external recipients, let the send continue
        if (externalRecipients.length === 0) {
          event.completed({ allowEvent: true });
          return;
        }
        
        // Get attachments if any
        item.getAttachmentsAsync((attachmentsResult) => {
          if (attachmentsResult.status === Office.AsyncResultStatus.Failed) {
            console.error(`Error getting attachments: ${attachmentsResult.error.message}`);
            // Continue with just external recipients
            showExternalRecipientsDialog(externalRecipients, [], event);
            return;
          }

          const attachments = attachmentsResult.value || [];
          // Show dialog with external recipients and attachments
          showExternalRecipientsDialog(externalRecipients, attachments, event);
        });
      });
    });
  });
}

// Create a type for the dialog data
interface ExternalRecipientData {
  displayName: string;
  emailAddress: string;
  selected: boolean;
}

interface AttachmentData {
  name: string;
  id: string;
  selected: boolean;
}

interface DialogData {
  externalRecipients: ExternalRecipientData[];
  attachments: AttachmentData[];
}

// Type guard to check if the object has a message property
function hasMessage(obj: any): obj is { message: string } {
  return obj && typeof obj.message === 'string';
}

/**
 * Shows a dialog with external recipients and attachments for confirmation
 */
function showExternalRecipientsDialog(
  externalRecipients: Office.EmailAddressDetails[], 
  attachments: Office.AttachmentDetailsCompose[], 
  event: Office.AddinCommands.Event
) {
  // Store the data for the dialog to access
  Office.context.mailbox.item.saveAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.error(`Error saving item data: ${result.error.message}`);
      event.completed({ allowEvent: true });
      return;
    }

    // Create and store data for the dialog
    const dialogData: DialogData = {
      externalRecipients: externalRecipients.map(r => ({
        displayName: r.displayName,
        emailAddress: r.emailAddress,
        selected: true
      })),
      attachments: attachments.map(a => ({
        name: a.name,
        id: a.id,
        selected: true
      }))
    };
    
    // Store the data for the dialog to access later
    if (typeof window !== 'undefined' && window.sessionStorage) {
      window.sessionStorage.setItem('externalRecipientsData', JSON.stringify(dialogData));
    }
    
    // Get the base URL for the add-in
    const baseUrl = getBaseUrl();
    
    // Encode the data for URL parameters (useful for OWA where sessionStorage may not work)
    const encodedData = encodeURIComponent(JSON.stringify(dialogData));
    const dialogUrl = `${baseUrl}confirm-dialog.html?data=${encodedData}`;
    
    // Determine if we're in OWA by checking for browser features
    const isOwa = isRunningInOwa();
    
    // Display the dialog with appropriate settings for the environment
    Office.context.ui.displayDialogAsync(
      dialogUrl,
      { 
        height: 75,  // Increased height for better content display in OWA
        width: 45,   // Slightly wider for better readability
        displayInIframe: !isOwa, // Use popup for OWA, iframe for desktop
        promptBeforeOpen: false
      },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error(`Error displaying dialog: ${result.error.message}`);
          console.error(`Dialog error code: ${result.error.code}`);
          
          // Show a fallback notification and let the send continue
          showNotification("Cannot display confirmation dialog. Please check external recipients carefully.");
          event.completed({ allowEvent: true });
          return;
        }

        // Get the dialog instance
        const dialog = result.value;
        
        // Handle dialog events
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (messageArg) => {
          try {
            const messageObj = messageArg as any;
            let messageText = "";
            
            // Safely extract the message text
            if (hasMessage(messageObj)) {
              messageText = messageObj.message;
            } else if (typeof messageObj === 'string') {
              messageText = messageObj;
            }
            
            // Try to parse the message
            const messageFromDialog = JSON.parse(messageText);
            
            if (messageFromDialog.action === 'send') {
              // User confirmed sending - set delay delivery time to 15 minutes from now
              const delayTime = new Date();
              delayTime.setMinutes(delayTime.getMinutes() + 15);

              // Check if delayDeliveryTime API is available
              if (Office.context.mailbox.item && 'delayDeliveryTime' in Office.context.mailbox.item) {
                // Set delayed delivery time using the correct API
                (Office.context.mailbox.item as any).delayDeliveryTime.setAsync(delayTime, (asyncResult) => {
                  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.error(`Failed to set delay: ${asyncResult.error.message}`);
                    showNotification("Could not set the 15-minute delay, but email will be sent.");
                    dialog.close();
                    event.completed({ allowEvent: true });
                    return;
                  }
                  
                  dialog.close();
                  event.completed({ allowEvent: true });
                });
              } else {
                // If the API is not available, just send without delay
                console.warn("delayDeliveryTime API not available in this Outlook version");
                showNotification("15-minute delay not available in this version of Outlook. Email will be sent immediately.");
                dialog.close();
                event.completed({ allowEvent: true });
              }
            } else if (messageFromDialog.action === 'cancel') {
              // User canceled sending
              dialog.close();
              event.completed({ allowEvent: false });
            }
          } catch (e) {
            console.error("Error parsing message from dialog", e);
            dialog.close();
            event.completed({ allowEvent: false });
          }
        });

        // Handle dialog closed
        dialog.addEventHandler(Office.EventType.DialogEventReceived, (_) => {
          // If dialog is closed without a message, cancel the send
          event.completed({ allowEvent: false });
        });
      }
    );
  });
}

/**
 * Shows a notification to the user
 */
function showNotification(message: string) {
  const notificationMessage: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: message,
    icon: "Icon.80x80",
    persistent: true,
  };

  Office.context.mailbox.item?.notificationMessages.replaceAsync(
    "ExternalRecipientsNotification",
    notificationMessage
  );
}

/**
 * Gets the base URL of the add-in
 */
function getBaseUrl(): string {
  // Try to determine the base URL from the current document
  try {
    // This works for commands.html
    if (typeof window !== 'undefined' && window.location) {
      const url = new URL(window.location.href);
      const pathParts = url.pathname.split('/');
      // Remove the last part (commands.html)
      pathParts.pop();
      return url.origin + pathParts.join('/') + '/';
    }
  } catch (e) {
    console.error("Error determining base URL:", e);
  }
  
  // Default to current origin or fallback for OWA
  return typeof window !== 'undefined' && window.location ? 
    window.location.origin + '/' : 
    'https://localhost:3000/';
}

/**
 * Determines if the add-in is running in Outlook Web Access
 */
function isRunningInOwa(): boolean {
  if (Office.context.mailbox) {
    // Use the host information from Office.context
    const hostName = Office.context.mailbox.diagnostics?.hostName || '';
    return hostName.indexOf('OWA') > -1 || hostName.indexOf('WebApp') > -1;
  }
  
  // Fallback detection method
  return typeof window !== 'undefined' && 
         window.navigator && 
         window.navigator.userAgent &&
         (window.navigator.userAgent.indexOf('Outlook Web App') > -1 ||
          window.navigator.userAgent.indexOf('OWA') > -1);
}

// Register the functions with Office
Office.actions.associate("action", action);
Office.actions.associate("validateExternalRecipients", validateExternalRecipients);
