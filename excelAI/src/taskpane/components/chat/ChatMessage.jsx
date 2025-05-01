import React from "react";
import {
  makeStyles,
  tokens,
  Card,
  CardHeader,
  CardPreview,
  Avatar,
  Text,
  Tooltip
} from "@fluentui/react-components";
import { 
  Bot24Regular,
  Person24Regular,
  ArrowRotateClockwise20Regular
} from "@fluentui/react-icons";

const useStyles = makeStyles({
  message: {
    maxWidth: "92%",
    marginBottom: "16px",
    animation: "fadeIn 0.3s ease-in-out",
  },
  userMessage: {
    alignSelf: "flex-end",
    marginLeft: "auto",
    position: "relative", // For positioning the resend button
  },
  aiMessage: {
    alignSelf: "flex-start",
    marginRight: "auto",
  },
  messageContent: {
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
    padding: "4px 0 4px 0",
    fontSize: tokens.fontSizeBase300,
    lineHeight: tokens.lineHeightBase300,
  },
  aiCard: {
    backgroundColor: tokens.colorBrandBackground2,
    borderLeft: `4px solid ${tokens.colorBrandStroke1}`,
    padding: "16px 20px",
    boxShadow: tokens.shadow4,
  },
  userCard: {
    backgroundColor: tokens.colorNeutralBackground3,
    padding: "16px 20px",
    boxShadow: tokens.shadow4,
  },
  resendButton: {
    position: "absolute",
    top: "-8px",
    left: "-8px",
    zIndex: 10,
    backgroundColor: tokens.colorNeutralBackground1,
    boxShadow: tokens.shadow4,
    borderRadius: "50%",
    width: "28px",
    height: "28px",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    cursor: "pointer",
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    transition: "transform 0.2s ease",
    "&:hover": {
      transform: "scale(1.1)",
      backgroundColor: tokens.colorNeutralBackground3,
    }
  },
  avatar: {
    boxShadow: tokens.shadow4,
  },
  codeBlock: {
    backgroundColor: tokens.colorNeutralBackground3,
    padding: "16px",
    borderRadius: tokens.borderRadiusMedium,
    fontFamily: "Consolas, Monaco, 'Andale Mono', monospace",
    overflowX: "auto",
    marginTop: "12px",
    marginBottom: "12px",
    fontSize: tokens.fontSizeBase200,
    lineHeight: tokens.lineHeightBase200,
    boxShadow: "inset 0 0 6px rgba(0,0,0,0.1)",
  },
  successMessage: {
    backgroundColor: tokens.colorPaletteGreenBackground1,
    borderLeft: `4px solid ${tokens.colorPaletteGreenBorderActive}`,
  },
  errorMessage: {
    backgroundColor: tokens.colorPaletteRedBackground1,
    borderLeft: `4px solid ${tokens.colorPaletteRedBorderActive}`,
  }
});

const ChatMessage = ({ message, onResend }) => {
  const styles = useStyles();
  
  // Render the message content
  const renderMessageContent = () => {
    // Check if the message contains code blocks (marked with backticks)
    const codeBlockRegex = /```(js|javascript|vba)?\s*([\s\S]*?)```/g;
    
    if (message.content.match(codeBlockRegex)) {
      // Split the content by code blocks
      const parts = [];
      let lastIndex = 0;
      let match;
      
      // Reset regex index
      codeBlockRegex.lastIndex = 0;
      
      while ((match = codeBlockRegex.exec(message.content)) !== null) {
        // Add text before code block
        if (match.index > lastIndex) {
          parts.push({
            type: 'text',
            content: message.content.substring(lastIndex, match.index)
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
      if (lastIndex < message.content.length) {
        parts.push({
          type: 'text',
          content: message.content.substring(lastIndex)
        });
      }
      
      // Render each part
      return (
        <>
          {parts.map((part, index) => (
            part.type === 'text' ? (
              <Text key={index} className={styles.messageContent}>{part.content}</Text>
            ) : (
              <div key={index} className={styles.codeBlock}>
                <Text>{part.content}</Text>
              </div>
            )
          ))}
        </>
      );
    }
    
    // Default: Render plain text message
    return (
      <Text className={styles.messageContent}>{message.content}</Text>
    );
  };
  
  return (
    <div 
      className={`${styles.message} ${
        message.role === "user" ? styles.userMessage : styles.aiMessage
      }`}
    >
      {message.role === "user" && (
        <Tooltip content="Resend this message" relationship="label">
          <div 
            className={styles.resendButton}
            onClick={() => onResend(message.content)}
            aria-label="Resend message"
          >
            <ArrowRotateClockwise20Regular />
          </div>
        </Tooltip>
      )}
      <Card 
        className={
          message.role === "user" 
            ? styles.userCard 
            : message.isSuccess 
              ? `${styles.aiCard} ${styles.successMessage}`
              : message.isError
                ? `${styles.aiCard} ${styles.errorMessage}`
                : styles.aiCard
        }
      >
        <CardHeader
          image={
            <Avatar 
              className={styles.avatar}
              icon={message.role === "user" ? <Person24Regular /> : <Bot24Regular />} 
              color={message.role === "user" ? "neutral" : message.isError ? "danger" : message.isSuccess ? "success" : "brand"}
            />
          }
          header={
            <Text weight="semibold">
              {message.role === "user" ? "You" : "Excel AI Assistant"}
            </Text>
          }
        />
        <CardPreview>
          {renderMessageContent()}
        </CardPreview>
      </Card>
    </div>
  );
};

export default ChatMessage; 