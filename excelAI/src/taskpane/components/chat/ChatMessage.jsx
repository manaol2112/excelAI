import React from "react";
import {
  makeStyles,
  tokens,
  Card,
  Avatar,
  Text,
  Tooltip,
  Button,
  mergeClasses,
  Divider
} from "@fluentui/react-components";
import { 
  Bot24Regular,
  Person24Regular,
  ArrowRotateClockwise20Regular,
  CopyRegular,
  MoreHorizontalRegular,
  CheckmarkRegular,
  ErrorCircleRegular,
  InfoRegular,
  Table24Regular
} from "@fluentui/react-icons";

const useStyles = makeStyles({
  message: {
    maxWidth: "92%",
    marginBottom: "16px",
    animation: "messageBounceIn 0.4s cubic-bezier(0.18, 1.25, 0.4, 1)",
    position: "relative",
  },
  userMessage: {
    alignSelf: "flex-end",
    marginLeft: "auto",
  },
  aiMessage: {
    alignSelf: "flex-start",
    marginRight: "auto",
  },
  messageContent: {
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
    padding: "2px 0 2px 0",
    fontSize: tokens.fontSizeBase300,
    lineHeight: tokens.lineHeightBase500,
    color: tokens.colorNeutralForeground1,
  },
  aiCard: {
    backgroundColor: tokens.colorNeutralBackground1,
    padding: "14px 18px",
    boxShadow: "0 1px 4px rgba(0, 0, 0, 0.06)",
    borderRadius: "18px 18px 18px 4px",
    border: `1px solid ${tokens.colorNeutralStrokeAccessible}`,
    transition: "all 0.2s ease",
    "&:hover": {
      boxShadow: "0 3px 10px rgba(0, 0, 0, 0.08)",
    }
  },
  userCard: {
    backgroundColor: "#F0F8FF",
    padding: "14px 18px",
    boxShadow: "0 1px 4px rgba(0, 0, 0, 0.08)",
    borderRadius: "18px 18px 4px 18px",
    border: `1px solid #E8F4FF`,
    transition: "all 0.2s ease",
    "&:hover": {
      boxShadow: "0 3px 10px rgba(0, 0, 0, 0.1)",
    }
  },
  resendButton: {
    position: "absolute",
    top: "-8px",
    left: "-8px",
    zIndex: 10,
    backgroundColor: tokens.colorNeutralBackground1,
    boxShadow: "0 2px 6px rgba(0, 0, 0, 0.1)",
    borderRadius: "50%",
    width: "24px",
    height: "24px",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    cursor: "pointer",
    transition: "all 0.2s ease",
    "&:hover": {
      transform: "scale(1.15)",
      backgroundColor: tokens.colorNeutralBackground2,
    }
  },
  avatar: {
    boxShadow: "0 1px 3px rgba(0, 0, 0, 0.08)",
  },
  userAvatar: {
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralBackground1,
    border: "none",
  },
  aiAvatar: {
    backgroundColor: tokens.colorNeutralBackground1,
    color: tokens.colorBrandForeground1,
    border: `1px solid ${tokens.colorNeutralStrokeAccessible}`,
  },
  codeBlock: {
    backgroundColor: tokens.colorNeutralBackground3,
    padding: "12px 16px",
    borderRadius: "8px",
    fontFamily: "Consolas, Monaco, 'Andale Mono', monospace",
    overflowX: "auto",
    marginTop: "12px",
    marginBottom: "12px",
    fontSize: tokens.fontSizeBase200,
    lineHeight: tokens.lineHeightBase300,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    position: "relative",
  },
  codeHeader: {
    position: "absolute",
    top: "4px",
    right: "4px",
    display: "flex",
    gap: "2px",
    padding: "2px",
  },
  copyButton: {
    padding: "2px",
    minWidth: "auto",
    height: "auto",
    color: tokens.colorNeutralForeground3,
    "&:hover": {
      color: tokens.colorNeutralForeground1,
      backgroundColor: "transparent",
    }
  },
  successMessage: {
    borderLeft: `3px solid ${tokens.colorStatusSuccessForeground1}`,
  },
  errorMessage: {
    borderLeft: `3px solid ${tokens.colorStatusDangerForeground1}`,
  },
  thinkingContainer: {
    display: "flex",
    justifyContent: "flex-start",
    alignItems: "center", 
    padding: "8px 0",
    animation: "fadeIn 0.3s ease-out forwards",
  },
  dotLoader: {
    display: 'flex',
    alignItems: 'flex-end',
    justifyContent: 'center',
    height: '1.5em',
    minWidth: '3em',
    gap: '0.4em',
    marginLeft: '8px',
  },
  dot: {
    display: 'inline-block',
    width: '0.6em',
    height: '0.6em',
    margin: 0,
    borderRadius: '50%',
    background: '#0078d4',
    opacity: 0.9,
    animation: 'bounce 1.4s infinite ease-in-out',
    transform: 'translateY(0px)',
  },
  dot1: { 
    animationDelay: '0s' 
  },
  dot2: { 
    animationDelay: '0.32s' 
  },
  dot3: { 
    animationDelay: '0.64s' 
  },
  '@keyframes bounce': {
    '0%': { transform: 'translateY(0px)' },
    '30%': { transform: 'translateY(-10px)' },
    '60%': { transform: 'translateY(0px)' },
    '100%': { transform: 'translateY(0px)' }
  },
  messageMeta: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    marginBottom: "6px",
  },
  messageHeader: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
  },
  messageTime: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
    opacity: 0.9,
  },
  statusIcon: {
    marginRight: "6px",
    flexShrink: 0,
  },
  userText: {
    color: tokens.colorNeutralForeground1,
  },
  messageActions: {
    display: "flex",
    justifyContent: "flex-start",
    gap: "4px",
    marginTop: "6px",
    opacity: 0,
    transition: "opacity 0.2s ease",
    "$message:hover &": {
      opacity: 0.8,
    }
  },
  actionButton: {
    minWidth: "28px",
    height: "28px",
    color: tokens.colorNeutralForeground3,
    "&:hover": {
      color: tokens.colorNeutralForeground1,
      backgroundColor: "transparent",
    }
  },
  "@keyframes messageBounceIn": {
    "0%": { 
      opacity: 0, 
      transform: "translateY(8px) scale(0.98)" 
    },
    "70%": { 
      opacity: 1, 
      transform: "translateY(-2px) scale(1.01)" 
    },
    "100%": { 
      transform: "translateY(0) scale(1)" 
    },
  },
  "@keyframes fadeIn": {
    "0%": { opacity: 0 },
    "100%": { opacity: 1 }
  },
  dataCard: {
    backgroundColor: tokens.colorNeutralBackground2,
    padding: "14px 18px",
    boxShadow: "0 1px 4px rgba(0, 0, 0, 0.06)",
    borderRadius: "18px 18px 18px 4px",
    border: `1px solid ${tokens.colorNeutralStrokeAccessible}`,
    borderLeft: `3px solid ${tokens.colorBrandBackground}`,
    transition: "all 0.2s ease",
    "&:hover": {
      boxShadow: "0 3px 10px rgba(0, 0, 0, 0.08)",
    }
  },
  dataAvatar: {
    backgroundColor: tokens.colorNeutralBackground1,
    color: tokens.colorBrandBackground,
    border: `1px solid ${tokens.colorBrandBackground}`,
  },
  dataCodeBlock: {
    backgroundColor: tokens.colorNeutralBackground4,
    padding: "12px 16px",
    borderRadius: "8px",
    fontFamily: "Consolas, Monaco, 'Andale Mono', monospace",
    overflowX: "auto",
    marginTop: "12px",
    marginBottom: "12px",
    fontSize: tokens.fontSizeBase200,
    lineHeight: tokens.lineHeightBase300,
    border: `1px solid ${tokens.colorBrandBackground}`,
    borderLeft: `4px solid ${tokens.colorBrandBackground}`,
    position: "relative",
    maxHeight: "300px",
    overflowY: "auto"
  },
});

const ChatMessage = ({ message, onResend, style }) => {
  const styles = useStyles();
  const isUser = message && message.role === "user";
  const isData = message && message.isData === true;
  
  // Format timestamp
  const timestamp = message && message.timestamp ? new Date(message.timestamp).toLocaleTimeString([], { hour: '2-digit', minute:'2-digit' }) : '';
  
  // Ensure message content is a string
  const safeContent = message && typeof message.content === 'string' ? message.content : 
                      message && message.content ? String(message.content) : '';
  
  // Render thinking indicator
  const renderThinkingIndicator = () => {
    return (
      <div className={styles.thinkingContainer}>
        <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>AI is thinking</Text>
        <div className={styles.dotLoader}>
          <div className={mergeClasses(styles.dot, styles.dot1)}></div>
          <div className={mergeClasses(styles.dot, styles.dot2)}></div>
          <div className={mergeClasses(styles.dot, styles.dot3)}></div>
        </div>
      </div>
    );
  };
  
  // Render the message content
  const renderMessageContent = () => {
    // Handle undefined or null message
    if (!message) {
      return <Text>Error: Invalid message</Text>;
    }
    
    if (message.isThinking) {
      return renderThinkingIndicator();
    }
    
    // Check if message content exists and is a string
    if (!safeContent) {
      return <Text>No content</Text>;
    }
    
    // Check if the message contains code blocks (marked with backticks)
    const codeBlockRegex = /```(js|javascript|vba)?\s*([\s\S]*?)```/g;
    
    if (safeContent.match(codeBlockRegex)) {
      // Split the content by code blocks
      const parts = [];
      let lastIndex = 0;
      let match;
      
      // Reset regex index
      codeBlockRegex.lastIndex = 0;
      
      while ((match = codeBlockRegex.exec(safeContent)) !== null) {
        // Add text before code block
        if (match.index > lastIndex) {
          parts.push({
            type: 'text',
            content: safeContent.substring(lastIndex, match.index)
          });
        }
        
        // Add code block
        parts.push({
          type: 'code',
          language: match[1] || '',
          content: match[2].trim()
        });
        
        lastIndex = match.index + match[0].length;
      }
      
      // Add any remaining text after the last code block
      if (lastIndex < safeContent.length) {
        parts.push({
          type: 'text',
          content: safeContent.substring(lastIndex)
        });
      }
      
      // Render each part
      return (
        <>
          {parts.map((part, index) => (
            part.type === 'text' ? (
              <Text 
                key={index} 
                className={mergeClasses(
                  styles.messageContent, 
                  isUser && styles.userText
                )}
              >
                {part.content}
              </Text>
            ) : (
              <div key={index} className={isData ? styles.dataCodeBlock : styles.codeBlock}>
                <div className={styles.codeHeader}>
                  <Button 
                    className={styles.copyButton} 
                    appearance="subtle" 
                    icon={<CopyRegular />}
                    onClick={() => navigator.clipboard.writeText(part.content)}
                    aria-label="Copy code"
                  />
                </div>
                <Text block>{part.content}</Text>
              </div>
            )
          ))}
        </>
      );
    }
    
    // Default: Render plain text message
    return (
      <Text 
        className={mergeClasses(
          styles.messageContent, 
          isUser && styles.userText
        )}
      >
        {safeContent}
      </Text>
    );
  };
  
  // Determine message status icon
  const getStatusIcon = () => {
    if (!message) return null;
    
    if (message.isError) {
      return <ErrorCircleRegular className={styles.statusIcon} />;
    } else if (message.isSuccess) {
      return <CheckmarkRegular className={styles.statusIcon} />;
    } else if (!isUser && !message.isThinking) {
      return <InfoRegular className={styles.statusIcon} />;
    }
    return null;
  };
  
  // If message is invalid, render nothing
  if (!message) {
    return null;
  }
  
  return (
    <div 
      className={`${styles.message} ${
        isUser ? styles.userMessage : styles.aiMessage
      }`}
      style={style}
    >
      {isUser && onResend && (
        <Tooltip content="Resend this message" relationship="label">
          <div 
            className={styles.resendButton}
            onClick={() => {
              if (typeof safeContent === 'string') {
                onResend(safeContent);
              }
            }}
            aria-label="Resend message"
          >
            <ArrowRotateClockwise20Regular />
          </div>
        </Tooltip>
      )}
      <Card 
        className={
          isUser 
            ? styles.userCard 
            : isData
              ? styles.dataCard
              : message.isSuccess 
                ? `${styles.aiCard} ${styles.successMessage}`
                : message.isError
                  ? `${styles.aiCard} ${styles.errorMessage}`
                  : styles.aiCard
        }
      >
        <div className={styles.messageMeta}>
          <div className={styles.messageHeader}>
            <Avatar 
              className={mergeClasses(
                styles.avatar, 
                isUser 
                  ? styles.userAvatar 
                  : isData 
                    ? styles.dataAvatar 
                    : styles.aiAvatar
              )}
              icon={isUser ? <Person24Regular /> : isData ? <Table24Regular /> : <Bot24Regular />} 
              size="tiny"
            />
            <Text weight="semibold" style={{ color: isUser ? tokens.colorNeutralForeground1Inverse : tokens.colorNeutralForeground1 }}>
              {isUser ? "You" : isData ? "Data" : "AI Assistant"}
            </Text>
            {timestamp && (
              <Text 
                className={styles.messageTime} 
                style={{ color: isUser ? tokens.colorNeutralForeground4Inverse : tokens.colorNeutralForeground3 }}
              >
                {timestamp}
              </Text>
            )}
          </div>
          {getStatusIcon()}
        </div>
        
        {renderMessageContent()}
        
        {!message.isThinking && safeContent && (
          <div className={styles.messageActions}>
            {!isUser && (
              <>
                <Button 
                  className={styles.actionButton}
                  appearance="subtle"
                  icon={<CopyRegular />}
                  onClick={() => navigator.clipboard.writeText(safeContent)}
                  aria-label="Copy message"
                />
              </>
            )}
            <Button 
              className={styles.actionButton}
              appearance="subtle"
              icon={<MoreHorizontalRegular />}
              aria-label="More options"
            />
          </div>
        )}
      </Card>
    </div>
  );
};

export default ChatMessage; 