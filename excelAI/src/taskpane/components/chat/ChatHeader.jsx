import React from 'react';
import {
  makeStyles,
  Button,
  Menu,
  MenuTrigger,
  MenuList,
  MenuItemRadio,
  MenuPopover,
  Text,
  tokens,
  Tooltip,
} from "@fluentui/react-components";
import {
  ChatSparkle24Regular,
  Bot24Regular,
  CheckmarkCircle24Filled,
  Eraser24Regular,
  ArrowUndo24Regular,
  Settings24Regular,
  ChevronDown20Regular
} from "@fluentui/react-icons";
import { AI_MODES } from '../AIChat'; // Assuming AI_MODES is exported from AIChat.jsx

const useStyles = makeStyles({
  header: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    padding: `${tokens.spacingVerticalL} ${tokens.spacingHorizontalL}`,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
    height: "68px", // Fixed height for the header
    boxSizing: "border-box"
  },
  modeSelector: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalS,
  },
  modeButton: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    borderRadius: tokens.borderRadiusMedium,
    cursor: 'pointer',
    border: `1px solid transparent`, // Add transparent border to prevent layout shifts
    transition: 'all 0.2s ease-in-out',
    backgroundColor: 'transparent',
    color: tokens.colorNeutralForeground2, // Default text color
    position: 'relative', // Needed for potential future indicators if required
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground1Hover,
      color: tokens.colorNeutralForeground1,
    },
    '& svg': {
      fontSize: '18px', // Slightly smaller icon
    }
  },
  selectedMode: {
    backgroundColor: tokens.colorBrandBackground2, // Subtle brand background
    color: tokens.colorBrandForeground1, // Brand text color
    fontWeight: tokens.fontWeightSemibold, // Make text bolder
    borderBottom: `2px solid ${tokens.colorBrandStroke1}`, // Distinct bottom border
    borderRadius: `${tokens.borderRadiusMedium} ${tokens.borderRadiusMedium} 0 0`, // Adjust border radius for bottom border
    ':hover': {
      backgroundColor: tokens.colorBrandBackground2Hover, // Slightly darker on hover
      color: tokens.colorBrandForeground1, // Keep brand color on hover
    },
  },
  modeText: {
    fontSize: tokens.fontSizeBase300,
    lineHeight: tokens.lineHeightBase300,
  },
  modeIcon: {
    marginRight: tokens.spacingHorizontalXS,
  },
  controls: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalS,
  },
  controlButton: {
    minWidth: "32px",
    maxWidth: "32px",
    height: "32px",
    padding: "0",
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
  },
  menuTriggerButton: {
    display: 'flex',
    alignItems: 'center',
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`, // Match mode button padding
    backgroundColor: 'transparent',
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: tokens.borderRadiusMedium,
    cursor: 'pointer',
    color: tokens.colorNeutralForeground1,
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground1Hover,
    },
    '& svg': {
      fontSize: '16px',
      marginLeft: tokens.spacingHorizontalXS,
    }
  },
});

const MODE_ICONS = {
  ASK: <ChatSparkle24Regular />,
  AGENT: <Bot24Regular />,
  PROMPT: <CheckmarkCircle24Filled />, // Example icon, replace if needed
};

const ChatHeader = ({ mode, onModeChange, onClearChat, canUndo, onUndo, onSettingsClick }) => {
  const styles = useStyles();

  // Simplified mode selection using buttons for a tab-like appearance
  return (
    <div className={styles.header}>
      <div className={styles.modeSelector}>
        {Object.values(AI_MODES).map((modeOption) => (
          <Tooltip content={modeOption.description} relationship="label" withArrow>
            <button
              key={modeOption.key}
              className={`${styles.modeButton} ${mode.key === modeOption.key ? styles.selectedMode : ''}`}
              onClick={() => onModeChange(modeOption)}
              aria-pressed={mode.key === modeOption.key}
            >
              {React.cloneElement(MODE_ICONS[modeOption.key], { className: styles.modeIcon })}
              <Text className={styles.modeText}>{modeOption.text}</Text>
            </button>
          </Tooltip>
        ))}
      </div>

      <div className={styles.controls}>
        <Tooltip content="Undo last AI action" relationship="label" withArrow>
            <Button
                appearance="subtle"
                icon={<ArrowUndo24Regular />}
                onClick={onUndo}
                disabled={!canUndo}
                className={styles.controlButton}
                aria-label="Undo last action"
            />
        </Tooltip>
        <Tooltip content="Clear chat history" relationship="label" withArrow>
            <Button
                appearance="subtle"
                icon={<Eraser24Regular />}
                onClick={onClearChat}
                className={styles.controlButton}
                aria-label="Clear chat"
            />
        </Tooltip>
        {/* <Tooltip content="Settings" relationship="label" withArrow>
            <Button
                appearance="subtle"
                icon={<Settings24Regular />}
                onClick={onSettingsClick} // Assuming you have an onSettingsClick handler
                className={styles.controlButton}
                aria-label="Settings"
            />
        </Tooltip> */}
      </div>
    </div>
  );
};

export default ChatHeader;