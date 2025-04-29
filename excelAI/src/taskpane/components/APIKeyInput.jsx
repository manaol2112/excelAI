import * as React from "react";
import { useState } from "react";
import {
  Input,
  Button,
  Dialog,
  DialogTrigger,
  DialogSurface,
  DialogTitle,
  DialogContent,
  DialogBody,
  DialogActions,
  Text,
  Field,
  Link,
  makeStyles,
  tokens,
  useId,
  Card,
  CardHeader
} from "@fluentui/react-components";
import { Key24Regular, Info24Regular, CheckmarkCircle24Regular, DismissCircle24Regular, ShieldLock24Regular } from "@fluentui/react-icons";
import { useAI } from "../../context/AIContext";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    padding: "16px",
    background: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    boxShadow: tokens.shadow4,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  header: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    paddingBottom: "12px",
  },
  title: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase400,
  },
  keyInput: {
    width: "100%",
  },
  infoText: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    marginBottom: "12px",
    lineHeight: tokens.lineHeightBase200,
  },
  securityNote: {
    display: "flex",
    alignItems: "flex-start",
    gap: "8px",
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusMedium,
    marginTop: "12px",
  },
  securityIcon: {
    color: tokens.colorBrandForeground1,
    flexShrink: 0,
  },
  dialogContent: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
  },
  statusIndicator: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    padding: "4px 8px",
    borderRadius: tokens.borderRadiusMedium,
  },
  validKey: {
    color: tokens.colorStatusSuccessForeground1,
    backgroundColor: tokens.colorStatusSuccessBackground1,
  },
  invalidKey: {
    color: tokens.colorStatusDangerForeground1,
    backgroundColor: tokens.colorStatusDangerBackground1,
  },
  actionButton: {
    marginTop: "12px",
    display: "flex",
    justifyContent: "center",
  }
});

const APIKeyInput = () => {
  const { apiKey, setApiKey, isApiKeyValid } = useAI();
  const [tempApiKey, setTempApiKey] = useState(apiKey);
  const [isDialogOpen, setIsDialogOpen] = useState(false);
  const styles = useStyles();
  const inputId = useId("api-key");

  const handleSaveApiKey = () => {
    setApiKey(tempApiKey);
    setIsDialogOpen(false);
  };

  const handleClearApiKey = () => {
    setTempApiKey("");
    setApiKey("");
    setIsDialogOpen(false);
  };

  return (
    <Card className={styles.container}>
      <CardHeader 
        header={
          <div className={styles.header}>
            <div className={styles.title}>
              <Key24Regular />
              <Text>OpenAI API Key</Text>
            </div>
            {isApiKeyValid ? (
              <div className={`${styles.statusIndicator} ${styles.validKey}`}>
                <CheckmarkCircle24Regular />
                <Text size={200}>Valid</Text>
              </div>
            ) : (
              <div className={`${styles.statusIndicator} ${styles.invalidKey}`}>
                <DismissCircle24Regular />
                <Text size={200}>Not Set</Text>
              </div>
            )}
          </div>
        }
      />

      <Text size={300}>
        {isApiKeyValid 
          ? "Your API key is set and ready to use with Excel AI." 
          : "You need to set your OpenAI API key to use AI features."
        }
      </Text>

      <div className={styles.actionButton}>
        <Dialog open={isDialogOpen} onOpenChange={(e, data) => setIsDialogOpen(data.open)}>
          <DialogTrigger disableButtonEnhancement>
            <Button appearance="primary" icon={<Key24Regular />}>
              {apiKey ? "Change API Key" : "Set API Key"}
            </Button>
          </DialogTrigger>
          <DialogSurface>
            <DialogBody>
              <DialogTitle>OpenAI API Key</DialogTitle>
              <DialogContent className={styles.dialogContent}>
                <Text className={styles.infoText}>
                  Enter your OpenAI API key to use the AI features in Excel AI.
                </Text>
                <Field
                  label="API Key"
                  id={inputId}
                  hint={
                    <Link href="https://platform.openai.com/api-keys" target="_blank">
                      Get your API key from OpenAI
                    </Link>
                  }
                >
                  <Input
                    id={inputId}
                    className={styles.keyInput}
                    type="password"
                    value={tempApiKey}
                    onChange={(e) => setTempApiKey(e.target.value)}
                    contentBefore={<Key24Regular />}
                    size="large"
                  />
                </Field>
                
                <div className={styles.securityNote}>
                  <ShieldLock24Regular className={styles.securityIcon} />
                  <Text size={200}>
                    Your API key is stored locally in your browser and is never sent to our servers. 
                    It's used directly from your device to communicate with OpenAI.
                  </Text>
                </div>
              </DialogContent>
              <DialogActions>
                <Button appearance="secondary" onClick={handleClearApiKey}>
                  Clear
                </Button>
                <Button appearance="primary" onClick={handleSaveApiKey} disabled={!tempApiKey}>
                  Save
                </Button>
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>
      </div>
    </Card>
  );
};

export default APIKeyInput; 