import React, { useState } from 'react';
import {
  makeStyles,
  tokens,
  Input,
  Button,
  Textarea,
  mergeClasses
} from "@fluentui/react-components";
import { Send24Regular, Mic24Regular, Attach24Regular } from "@fluentui/react-icons";

const useStyles = makeStyles({
  container: {
    padding: "16px",
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
    boxShadow: `0 -2px 10px rgba(0, 0, 0, 0.05)`,
    position: "relative",
    zIndex: 1,
  },
  inputContainer: {
    display: "flex",
    flexDirection: "column",
    width: "100%",
    position: "relative",
    borderRadius: "16px",
    backgroundColor: tokens.colorNeutralBackground1,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    boxShadow: tokens.shadow4,
    transition: "all 0.2s ease",
    padding: "4px",
    "&:hover": {
      borderColor: tokens.colorNeutralStroke1,
      boxShadow: tokens.shadow8,
    },
    "&:focus-within": {
      borderColor: tokens.colorBrandStroke1,
      boxShadow: `0 0 0 2px ${tokens.colorBrandStroke1Hover}`,
    }
  },
  textarea: {
    width: "100%",
    resize: "none",
    border: "none",
    backgroundColor: "transparent",
    padding: "8px 12px",
    fontFamily: tokens.fontFamilyBase,
    fontSize: tokens.fontSizeBase300,
    lineHeight: tokens.lineHeightBase300,
    "&:focus": {
      outline: "none",
    },
    maxHeight: "120px",
  },
  buttonsContainer: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    padding: "4px 8px",
  },
  auxiliaryButtons: {
    display: "flex",
    gap: "4px",
  },
  actionButton: {
    color: tokens.colorNeutralForeground3,
    width: "32px",
    height: "32px",
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
    borderRadius: "50%",
    transition: "all 0.2s ease",
    "&:hover": {
      backgroundColor: tokens.colorNeutralBackground3,
      color: tokens.colorNeutralForeground1,
    }
  },
  sendButton: {
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
    width: "36px",
    height: "36px",
    borderRadius: "50%",
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
    cursor: "pointer",
    transition: "all 0.2s ease",
    border: "none",
    "&:hover": {
      backgroundColor: tokens.colorBrandBackgroundHover,
      transform: "scale(1.05)",
    },
    "&:disabled": {
      backgroundColor: tokens.colorNeutralBackground4,
      color: tokens.colorNeutralForeground3,
      cursor: "not-allowed",
      transform: "none",
    }
  },
  placeholderText: {
    position: "absolute",
    top: "16px",
    left: "16px",
    color: tokens.colorNeutralForeground3,
    pointerEvents: "none",
    transition: "opacity 0.2s ease",
    opacity: 1,
  },
  hidePlaceholder: {
    opacity: 0,
  }
});

const ChatInput = ({ value, onChange, onKeyDown, onSend, disabled }) => {
  const styles = useStyles();
  const [isFocused, setIsFocused] = useState(false);
  
  // Safely determine if the button should be disabled
  const isSendDisabled = disabled || !value || (typeof value === 'string' && value.trim() === '');
  
  const handleFocus = () => setIsFocused(true);
  const handleBlur = () => setIsFocused(false);
  
  const handleInputChange = (e, data) => {
    if (onChange) {
      onChange(e, data.value);
    }
  };
  
  return (
    <div className={styles.container}>
      <div className={styles.inputContainer}>
        <Textarea
          className={styles.textarea}
          resize="none"
          value={value || ''}
          onChange={handleInputChange}
          onKeyDown={onKeyDown}
          disabled={disabled}
          onFocus={handleFocus}
          onBlur={handleBlur}
          rows={1}
          autoAdjustHeight
        />
        
        <div 
          className={mergeClasses(
            styles.placeholderText, 
            (isFocused || (value && value.length > 0)) && styles.hidePlaceholder
          )}
        >
          Ask about Excel formulas, data analysis, or help with your spreadsheet...
        </div>
        
        <div className={styles.buttonsContainer}>
          <div className={styles.auxiliaryButtons}>
            <Button 
              className={styles.actionButton}
              appearance="transparent"
              icon={<Attach24Regular />}
              disabled={disabled}
              aria-label="Attach file"
            />
            <Button
              className={styles.actionButton}
              appearance="transparent"
              icon={<Mic24Regular />}
              disabled={disabled}
              aria-label="Voice input"
            />
          </div>
          
          <Button
            className={styles.sendButton}
            appearance="transparent"
            icon={<Send24Regular />}
            onClick={onSend}
            disabled={isSendDisabled}
            aria-label="Send message"
          />
        </div>
      </div>
    </div>
  );
};

export default ChatInput; 