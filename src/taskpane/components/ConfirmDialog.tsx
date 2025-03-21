import * as React from "react";
import { useState } from "react";
import { 
  Button, 
  Checkbox, 
  Title1, 
  Subtitle1, 
  Divider, 
  makeStyles, 
  tokens,
  Card,
  Body1,
  Badge,
  Dialog,
  DialogTrigger,
  DialogSurface,
  DialogTitle,
  DialogContent,
  DialogBody,
  DialogActions,
  InfoLabel,
  Text
} from "@fluentui/react-components";
import { Warning24Regular, Timer24Regular } from "@fluentui/react-icons";

interface ExternalRecipient {
  displayName: string;
  emailAddress: string;
  selected: boolean;
}

interface Attachment {
  name: string;
  id: string;
  selected: boolean;
}

interface ConfirmDialogProps {
  externalRecipients: ExternalRecipient[];
  attachments: Attachment[];
  onSend: () => void;
  onCancel: () => void;
}

const useStyles = makeStyles({
  container: {
    padding: "20px",
    display: "flex",
    flexDirection: "column",
    gap: "20px",
    maxHeight: "100vh",
    overflow: "auto",
  },
  title: {
    display: "flex",
    alignItems: "center",
    gap: "10px",
    color: tokens.colorPaletteRedForeground1,
  },
  warningIcon: {
    color: tokens.colorPaletteRedForeground1,
  },
  section: {
    marginTop: "10px",
    marginBottom: "20px",
  },
  sectionTitle: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    marginBottom: "10px",
  },
  recipientsList: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    maxHeight: "200px",
    overflowY: "auto",
    padding: "10px",
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: tokens.borderRadiusMedium,
  },
  attachmentsList: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    maxHeight: "200px",
    overflowY: "auto",
    padding: "10px",
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: tokens.borderRadiusMedium,
  },
  recipientItem: {
    display: "flex",
    alignItems: "center",
    gap: "10px",
  },
  attachmentItem: {
    display: "flex",
    alignItems: "center",
    gap: "10px",
  },
  emailAddress: {
    color: tokens.colorNeutralForeground2,
    fontSize: tokens.fontSizeBase200,
  },
  actions: {
    display: "flex",
    justifyContent: "flex-end",
    gap: "10px",
    marginTop: "20px",
  },
  badge: {
    backgroundColor: tokens.colorPaletteRedBackground1,
    color: tokens.colorPaletteRedForeground1,
  },
  delayNotice: {
    backgroundColor: tokens.colorNeutralBackground2,
    padding: "12px",
    borderRadius: tokens.borderRadiusMedium,
    display: "flex",
    alignItems: "center",
    gap: "8px",
  },
  delayIcon: {
    color: tokens.colorBrandForeground1,
  },
  delayText: {
    color: tokens.colorNeutralForeground1,
  },
});

const ConfirmDialog: React.FC<ConfirmDialogProps> = (props) => {
  const { externalRecipients, attachments, onSend, onCancel } = props;
  const styles = useStyles();

  // State for recipients and attachments
  const [recipients, setRecipients] = useState<ExternalRecipient[]>(externalRecipients);
  const [attachmentList, setAttachmentList] = useState<Attachment[]>(attachments);
  
  // Handlers for toggling selection
  const toggleRecipient = (index: number) => {
    const updatedRecipients = [...recipients];
    updatedRecipients[index].selected = !updatedRecipients[index].selected;
    setRecipients(updatedRecipients);
  };

  const toggleAttachment = (index: number) => {
    const updatedAttachments = [...attachmentList];
    updatedAttachments[index].selected = !updatedAttachments[index].selected;
    setAttachmentList(updatedAttachments);
  };

  // Toggle all recipients
  const toggleAllRecipients = (checked: boolean) => {
    setRecipients(
      recipients.map(recipient => ({
        ...recipient,
        selected: checked
      }))
    );
  };

  // Toggle all attachments
  const toggleAllAttachments = (checked: boolean) => {
    setAttachmentList(
      attachmentList.map(attachment => ({
        ...attachment,
        selected: checked
      }))
    );
  };

  // Check if there are any selected recipients or attachments
  const hasSelectedRecipients = recipients.some(r => r.selected);

  return (
    <Dialog open={true} modalType="alert">
      <DialogSurface>
        <DialogBody>
          <DialogTitle>
            <div className={styles.title}>
              <Warning24Regular className={styles.warningIcon} />
              <Title1>Confirm External E-mail</Title1>
            </div>
          </DialogTitle>
          <DialogContent>
            <div className={styles.container}>
              <Body1>
                You're about to send an email to recipients outside your organization. 
                Please review the recipients and attachments carefully.
              </Body1>
              
              <div className={styles.delayNotice}>
                <Timer24Regular className={styles.delayIcon} />
                <div>
                  <Text weight="semibold">15-Minute Sending Delay</Text>
                  <Text className={styles.delayText} size={200}>
                    This email will be held for 15 minutes before sending to external recipients, giving you a chance to recall it if needed.
                  </Text>
                </div>
              </div>
              
              <div className={styles.section}>
                <div className={styles.sectionTitle}>
                  <Subtitle1>External Recipients</Subtitle1>
                  <Badge className={styles.badge} appearance="filled">{recipients.length}</Badge>
                </div>
                
                <Checkbox 
                  label="Select All" 
                  checked={recipients.every(r => r.selected)}
                  onChange={(_event, data) => toggleAllRecipients(!!data.checked)}
                />
                
                <div className={styles.recipientsList}>
                  {recipients.map((recipient, index) => (
                    <div key={index} className={styles.recipientItem}>
                      <Checkbox 
                        checked={recipient.selected}
                        onChange={() => toggleRecipient(index)}
                      />
                      <div>
                        <div>{recipient.displayName || recipient.emailAddress}</div>
                        {recipient.displayName && (
                          <div className={styles.emailAddress}>{recipient.emailAddress}</div>
                        )}
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              {attachmentList.length > 0 && (
                <div className={styles.section}>
                  <div className={styles.sectionTitle}>
                    <Subtitle1>Attachments</Subtitle1>
                    <Badge className={styles.badge} appearance="filled">{attachmentList.length}</Badge>
                  </div>
                  
                  <Checkbox 
                    label="Select All" 
                    checked={attachmentList.every(a => a.selected)}
                    onChange={(_event, data) => toggleAllAttachments(!!data.checked)}
                  />
                  
                  <div className={styles.attachmentsList}>
                    {attachmentList.map((attachment, index) => (
                      <div key={index} className={styles.attachmentItem}>
                        <Checkbox 
                          checked={attachment.selected}
                          onChange={() => toggleAttachment(index)}
                        />
                        <div>{attachment.name}</div>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={onCancel}>
              Cancel
            </Button>
            <Button 
              appearance="primary" 
              onClick={onSend}
              disabled={!hasSelectedRecipients}
            >
              Schedule Send
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};

export default ConfirmDialog;