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
    return !email.toLowerCase().endsWith("@kairos.com");
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
    
    // Display the dialog
    Office.context.ui.displayDialogAsync(
      'https://localhost:3000/confirm-dialog.html',
      { height: 60, width: 40, displayInIframe: true },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error(`Error displaying dialog: ${result.error.message}`);
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

              // Set delayed delivery time using the correct API
              (Office.context.mailbox.item as any).delayDeliveryTime.setAsync(delayTime, (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                  console.error(`Failed to set delay: ${asyncResult.error.message}`);
                  dialog.close();
                  event.completed({ allowEvent: false });
                  return;
                }
                
                dialog.close();
                event.completed({ allowEvent: true });
              });
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

// Register the functions with Office
Office.actions.associate("action", action);
Office.actions.associate("validateExternalRecipients", validateExternalRecipients);
