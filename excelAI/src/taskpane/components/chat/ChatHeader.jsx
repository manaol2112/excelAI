import React from 'react';
import {
  makeStyles,
  tokens,
  Text,
  Button,
  Menu,
  MenuTrigger,
  MenuPopover,
  MenuList,
  MenuItem,
  Badge,
  Tooltip,
  Divider
} from "@fluentui/react-components";
import { 
  Bot24Regular,
  ChevronDown20Regular,
  AppsAddIn24Regular,
  ArrowUndo24Regular,
  DeleteRegular,
  ChatRegular
} from "@fluentui/react-icons";

const useStyles = makeStyles({
  chatHeader: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: "8px",
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase500,
    padding: "12px 0",
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  headerLeft: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
  },
  headerRight: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
  },
  undoButton: {
    display: "flex",
    alignItems: "center",
    gap: "4px",
    fontSize: tokens.fontSizeBase200,
    padding: "4px 8px",
    height: "28px",
    backgroundColor: tokens.colorNeutralBackground2,
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground3,
    }
  },
  modeSelector: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    padding: "4px 8px",
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground2,
    cursor: "pointer",
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground3,
    }
  },
  clearButton: {
    display: "flex",
    alignItems: "center",
    gap: "4px",
    fontSize: tokens.fontSizeBase200,
    padding: "4px 8px",
    height: "28px",
    backgroundColor: tokens.colorNeutralBackground2,
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground3,
    }
  }
});

// AI operation modes
const AI_MODES = {
  ASK: { name: "Ask Mode", description: "Get answers in chat without modifying your spreadsheet" },
  AGENT: { name: "Agent Mode", description: "AI will directly apply changes to your spreadsheet" },
  PROMPT: { name: "Prompt Mode", description: "Ask permission before applying changes" }
};

const ChatHeader = ({ 
  mode, 
  onModeChange, 
  onClearChat, 
  canUndo, 
  onUndo
}) => {
  const styles = useStyles();
  
  return (
    <div className={styles.chatHeader}>
      <div className={styles.headerLeft}>
        <Bot24Regular />
        <Text>Excel AI Assistant</Text>
      </div>
      
      <div className={styles.headerRight}>
        <Tooltip content="Clear chat history" relationship="label">
          <Button 
            className={styles.clearButton}
            icon={<DeleteRegular />}
            onClick={onClearChat}
            appearance="subtle"
            size="small"
          >
            Clear Chat
          </Button>
        </Tooltip>
        
        <Divider vertical />
        
        <Tooltip content="Undo last operation" relationship="label">
          <Button 
            className={styles.undoButton}
            icon={<ArrowUndo24Regular />}
            onClick={onUndo}
            disabled={!canUndo}
            appearance="subtle"
            size="small"
          >
            Undo
          </Button>
        </Tooltip>
        
        <Divider vertical />
        
        <Menu>
          <MenuTrigger disableButtonEnhancement>
            <div className={styles.modeSelector}>
              <AppsAddIn24Regular />
              <Text>{AI_MODES[mode.key].name}</Text>
              <ChevronDown20Regular />
            </div>
          </MenuTrigger>
          <MenuPopover>
            <MenuList>
              <MenuItem 
                onClick={() => onModeChange(AI_MODES.ASK)}
                icon={mode.key === "ASK" ? <Badge appearance="filled" /> : null}
              >
                {AI_MODES.ASK.name}
                <Text size={100} block>{AI_MODES.ASK.description}</Text>
              </MenuItem>
              <MenuItem 
                onClick={() => onModeChange(AI_MODES.AGENT)}
                icon={mode.key === "AGENT" ? <Badge appearance="filled" /> : null}
              >
                {AI_MODES.AGENT.name}
                <Text size={100} block>{AI_MODES.AGENT.description}</Text>
              </MenuItem>
              <MenuItem 
                onClick={() => onModeChange(AI_MODES.PROMPT)}
                icon={mode.key === "PROMPT" ? <Badge appearance="filled" /> : null}
              >
                {AI_MODES.PROMPT.name}
                <Text size={100} block>{AI_MODES.PROMPT.description}</Text>
              </MenuItem>
            </MenuList>
          </MenuPopover>
        </Menu>
      </div>
    </div>
  );
};

export default ChatHeader; 