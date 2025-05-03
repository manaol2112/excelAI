import React, { useState, useRef, useEffect } from 'react';
import {
  makeStyles,
  tokens,
  Button,
  Textarea,
  mergeClasses,
  Tooltip,
  Divider,
  Badge
} from "@fluentui/react-components";
import { 
  Send24Regular, 
  Mic24Regular, 
  Attach24Regular,
  DocumentSearch24Regular,
  Table24Regular,
  TableSimple24Filled,
  DataBarVertical24Filled,
  FullScreenMaximize24Regular
} from "@fluentui/react-icons";

const useStyles = makeStyles({
  container: {
    padding: "20px",
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
    boxShadow: `0 -4px 16px rgba(0, 0, 0, 0.08)`,
    position: "relative",
    zIndex: 1,
  },
  inputContainer: {
    display: "flex",
    flexDirection: "column",
    width: "100%",
    position: "relative",
    borderRadius: "12px",
    backgroundColor: tokens.colorNeutralBackground1,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    boxShadow: tokens.shadow8,
    transition: "all 0.25s ease",
    "&:hover": {
      borderColor: tokens.colorNeutralStroke1,
      boxShadow: tokens.shadow16,
    },
    "&:focus-within": {
      borderColor: tokens.colorBrandStroke1,
      boxShadow: `0 0 0 2px ${tokens.colorBrandStroke1}`,
    }
  },
  textareaWrapper: {
    position: "relative",
    padding: "16px 18px 6px 18px",
  },
  textarea: {
    width: "100%",
    resize: "none",
    border: "none",
    backgroundColor: "transparent",
    fontFamily: tokens.fontFamilyBase,
    fontSize: tokens.fontSizeBase300,
    lineHeight: tokens.lineHeightBase400,
    color: tokens.colorNeutralForeground1,
    "&:focus": {
      outline: "none",
    },
    maxHeight: "120px",
    minHeight: "24px",
    padding: "0",
    margin: "0",
  },
  divider: {
    margin: "4px 0",
  },
  buttonsContainer: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    padding: "4px 8px 8px 8px",
  },
  auxiliaryButtons: {
    display: "flex",
    gap: "8px",
  },
  actionButton: {
    color: tokens.colorNeutralForeground3,
    width: "32px",
    height: "32px",
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
    borderRadius: "4px",
    transition: "all 0.2s ease",
    "&:hover": {
      backgroundColor: tokens.colorNeutralBackground3,
      color: tokens.colorBrandForeground1,
    }
  },
  selectionButton: {
    color: tokens.colorNeutralForeground3,
    width: "32px",
    height: "32px",
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
    borderRadius: "4px",
    transition: "all 0.2s ease",
    "&:hover": {
      backgroundColor: tokens.colorNeutralBackground3,
      color: tokens.colorBrandForeground1,
    }
  },
  activeSelectionButton: {
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForeground1BrandSelected,
    "&:hover": {
      backgroundColor: tokens.colorBrandBackgroundHover,
      color: tokens.colorNeutralForeground1BrandSelected,
    }
  },
  sendButton: {
    backgroundColor: "#0078d4",
    color: "#ffffff",
    width: "44px",
    height: "44px",
    borderRadius: "10px",
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
    cursor: "pointer",
    transition: "all 0.25s ease",
    border: "none",
    "&:hover": {
      backgroundColor: "#106ebe",
      transform: "translateY(-2px) scale(1.05)",
      boxShadow: tokens.shadow16,
    },
    "&:active": {
      transform: "translateY(1px)",
    },
    "&:disabled": {
      backgroundColor: tokens.colorNeutralBackground4,
      color: tokens.colorNeutralForeground3,
      cursor: "not-allowed",
      transform: "none",
      boxShadow: "none",
    }
  },
  placeholder: {
    position: "absolute",
    top: "12px",
    left: "16px",
    color: tokens.colorNeutralForeground3,
    pointerEvents: "none",
    transition: "all 0.2s ease",
    fontSize: tokens.fontSizeBase300,
    opacity: 1,
    display: "block",
  },
  placeholderHidden: {
    opacity: 0,
    display: "none",
  },
  placeholderFocused: {
    opacity: 0,
    display: "none",
  },
  characterCount: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    marginRight: "12px",
    userSelect: "none",
  },
  rightContainer: {
    display: "flex",
    alignItems: "center",
    gap: "12px",
  },
  selectionInfo: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorBrandForeground1,
    backgroundColor: tokens.colorBrandBackground,
    padding: "3px 8px",
    borderRadius: "4px",
    marginLeft: "8px",
    display: "flex",
    alignItems: "center",
    gap: "4px"
  },
  selectionBadge: {
    background: `linear-gradient(145deg, ${tokens.colorBrandBackground} 0%, ${tokens.colorBrandBackgroundHover} 100%)`,
    padding: "10px 14px",
    borderRadius: "8px",
    boxShadow: `${tokens.shadow4}, 0 0 0 1px ${tokens.colorBrandStroke1}`,
    marginBottom: "10px",
    transition: "all 0.2s ease",
    "&:hover": {
      boxShadow: `${tokens.shadow8}, 0 0 0 1px ${tokens.colorBrandStroke1}`,
      transform: "translateY(-1px)",
    }
  },
  selectionTitle: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralBackground1,
    marginBottom: "4px"
  },
  selectionDetails: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralBackground1,
    opacity: 0.9
  },
  selectionIndicator: {
    display: "inline-flex",
    alignItems: "center",
    gap: "4px",
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground1,
    padding: "4px 8px",
    borderRadius: "16px",
    fontSize: tokens.fontSizeBase100,
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
    maxWidth: "180px",
    margin: "0 8px",
    border: `1px solid ${tokens.colorBrandBackground}`,
    transition: "all 0.2s ease",
    animation: "pulseEffect 2s infinite ease-in-out",
    cursor: "pointer",
    "&:hover": {
      backgroundColor: tokens.colorBrandBackground,
      color: tokens.colorNeutralForeground1BrandSelected,
    }
  },
  dataIcon: {
    fontSize: tokens.fontSizeBase300,
    display: "flex",
    alignItems: "center",
  },
  autoSelectButton: {
    color: tokens.colorNeutralForeground3,
    width: "32px",
    height: "32px",
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
    borderRadius: "4px",
    transition: "all 0.2s ease",
    "&:hover": {
      backgroundColor: tokens.colorNeutralBackground3,
      color: tokens.colorBrandForeground1,
    }
  },
  "@keyframes pulseEffect": {
    "0%": { opacity: 0.8, transform: "scale(1)" },
    "50%": { opacity: 1, transform: "scale(1.05)" },
    "100%": { opacity: 0.8, transform: "scale(1)" }
  }
});

const MAX_CHARS = 10000;

const ChatInput = ({ value = '', onChange, onKeyDown, onSend, disabled, currentSelection, onToggleSelection, onAutoSelectRange }) => {
  const styles = useStyles();
  const [isFocused, setIsFocused] = useState(false);
  const textareaRef = useRef(null);
  const [charCount, setCharCount] = useState(0);
  
  useEffect(() => {
    // Update char count when value changes
    setCharCount(value ? value.length : 0);
  }, [value]);
  
  // Safely determine if the button should be disabled
  const isSendDisabled = disabled || !value || (typeof value === 'string' && value.trim() === '');
  const hasContent = value && typeof value === 'string' && value.trim().length > 0;
  
  const handleFocus = () => {
    setIsFocused(true);
  };
  
  const handleBlur = () => {
    setIsFocused(false);
  };
  
  const handleInputChange = (e, data) => {
    try {
      if (onChange) {
        // Ensure we're always passing the string value, not an object
        const actualValue = data && data.value !== undefined ? data.value : '';
        onChange(e, actualValue);
      }
    } catch (error) {
      console.error("Error in handleInputChange:", error);
    }
  };
  
  const handleSend = (e) => {
    try {
      if (!isSendDisabled && onSend) {
        onSend();
      }
    } catch (error) {
      console.error("Error in handleSend:", error);
    }
  };
  
  const handleToggleSelection = () => {
    if (onToggleSelection) {
      onToggleSelection();
    }
  };
  
  const handleAutoSelectRange = () => {
    if (onAutoSelectRange) {
      onAutoSelectRange();
    }
  };
  
  const focusInput = () => {
    if (textareaRef.current) {
      textareaRef.current.focus();
    }
  };
  
  // Calculate if we should show placeholder and how
  const showPlaceholder = false; // Always hide the placeholder
  const placeholderClasses = mergeClasses(
    styles.placeholder,
    !showPlaceholder && styles.placeholderHidden,
    isFocused && styles.placeholderFocused
  );
  
  // Determine if selection is active
  const hasActiveSelection = currentSelection && currentSelection.address;
  const selectionButtonClasses = mergeClasses(
    styles.selectionButton,
    hasActiveSelection && styles.activeSelectionButton
  );
  
  // Format selection address to be more concise
  const formatSelectionAddress = (address) => {
    if (!address) return '';
    
    // Show just the range part, not the sheet name
    const rangePart = address.includes('!') ? address.split('!')[1] : address;
    return rangePart;
  };
  
  return (
    <div className={styles.container} onClick={focusInput}>
      <div className={styles.inputContainer}>
        <div className={styles.textareaWrapper}>
          <Textarea
            ref={textareaRef}
            className={styles.textarea}
            resize="none"
            value={typeof value === 'string' ? value : ''}
            onChange={handleInputChange}
            onKeyDown={onKeyDown}
            disabled={disabled}
            onFocus={handleFocus}
            onBlur={handleBlur}
            rows={1}
            autoAdjustHeight
            maxLength={MAX_CHARS}
            placeholder="Ask about Excel formulas, data analysis, or help with your spreadsheet..."
          />
        </div>
        
        <Divider className={styles.divider} />
        
        <div className={styles.buttonsContainer}>
          <div className={styles.auxiliaryButtons}>
            <Tooltip content={hasActiveSelection ? "Data selection active - click to disable" : "Use Excel selection for analysis"} relationship="label">
              <Button 
                className={selectionButtonClasses}
                appearance="subtle"
                icon={hasActiveSelection ? <DataBarVertical24Filled /> : <TableSimple24Filled />}
                onClick={handleToggleSelection}
                aria-label={hasActiveSelection ? "Disable Excel selection" : "Use Excel selection"}
              />
            </Tooltip>
            
            <Tooltip content="Auto-select all data (used range)" relationship="label">
              <Button
                className={styles.autoSelectButton}
                appearance="subtle"
                icon={<FullScreenMaximize24Regular />}
                onClick={handleAutoSelectRange}
                aria-label="Auto-select used range"
              />
            </Tooltip>
            
            <Tooltip content="Voice input" relationship="label">
              <Button
                className={styles.actionButton}
                appearance="subtle"
                icon={<Mic24Regular />}
                disabled={disabled}
                aria-label="Voice input"
              />
            </Tooltip>
            
            {/* Remove or hide the document search button to save space when selection is active */}
            {!hasActiveSelection && (
              <Tooltip content="Search spreadsheet data" relationship="label">
                <Button
                  className={styles.actionButton}
                  appearance="subtle"
                  icon={<DocumentSearch24Regular />}
                  disabled={disabled}
                  aria-label="Search spreadsheet data"
                />
              </Tooltip>
            )}
          </div>
          
          <div className={styles.rightContainer}>
            {hasContent && (
              <span className={styles.characterCount}>{charCount}/{MAX_CHARS}</span>
            )}
            
            <Tooltip 
              content={isSendDisabled ? "Type a message to send" : "Send message"} 
              relationship="label"
            >
              <Button
                className={styles.sendButton}
                icon={<Send24Regular />}
                onClick={handleSend}
                disabled={isSendDisabled}
                aria-label="Send message"
              />
            </Tooltip>
          </div>
        </div>
      </div>
    </div>
  );
};

export default ChatInput; 