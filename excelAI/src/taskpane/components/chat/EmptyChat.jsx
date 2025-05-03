import React from "react";
import {
  makeStyles,
  tokens,
  Text,
  Button,
  mergeClasses
} from "@fluentui/react-components";
import { 
  Bot24Filled, 
  ArrowRight16Filled, 
  SparkleRegular, 
  DocumentSearchRegular, 
  CalculatorRegular,
  DataTrendingRegular
} from "@fluentui/react-icons";

const useStyles = makeStyles({
  emptyChatContainer: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    width: "100%",
    height: "100%",
    padding: "12px",
    gap: "12px",
  },
  header: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    marginBottom: "0",
  },
  botIcon: {
    fontSize: "28px",
    color: tokens.colorBrandForeground1,
    marginBottom: "8px",
  },
  titleContainer: {
    textAlign: "center",
    marginBottom: "8px",
  },
  title: {
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightSemibold,
    background: "linear-gradient(90deg, #0078D4, #2B88D8)",
    WebkitBackgroundClip: "text",
    WebkitTextFillColor: "transparent",
    marginBottom: "8px",
    lineHeight: "1.2",
  },
  subtitle: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    maxWidth: "320px",
    textAlign: "center",
    lineHeight: "1.3",
    marginBottom: "4px",
  },
  suggestionsContainer: {
    display: "flex",
    flexDirection: "column",
    width: "100%",
    gap: "8px",
    maxWidth: "400px",
    alignItems: "center",
  },
  suggestionButton: {
    width: "100%",
    padding: "10px 14px",
    borderRadius: "8px",
    cursor: "pointer",
    transition: "all 0.2s ease",
    textAlign: "left",
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    boxShadow: "0 1px 2px rgba(0,0,0,0.05)",
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    backgroundColor: tokens.colorNeutralBackground1,
    "&:hover": {
      transform: "translateY(-1px)",
      boxShadow: "0 2px 6px rgba(0,0,0,0.08)",
      borderColor: tokens.colorBrandStroke1,
      backgroundColor: tokens.colorNeutralBackground1Hover,
    }
  },
  suggestionContent: {
    display: "flex",
    alignItems: "center",
    gap: "10px",
  },
  suggestionIcon: {
    fontSize: "16px",
    color: tokens.colorBrandForeground1,
    flexShrink: 0,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
  },
  suggestionText: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightRegular,
    color: tokens.colorNeutralForeground1,
  },
  arrowIcon: {
    fontSize: "14px",
    color: tokens.colorNeutralForeground3,
    marginLeft: "8px",
  },
  footer: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
    maxWidth: "320px",
    textAlign: "center",
    marginTop: "4px",
  }
});

const EmptyChat = ({ mode, suggestions, onSuggestionClick }) => {
  const styles = useStyles();
  
  // Suggestion icons based on their index
  const getIcon = (index) => {
    switch(index) {
      case 0: return <DocumentSearchRegular className={styles.suggestionIcon} />;
      case 1: return <CalculatorRegular className={styles.suggestionIcon} />;
      case 2: return <DataTrendingRegular className={styles.suggestionIcon} />;
      case 3: return <SparkleRegular className={styles.suggestionIcon} />;
      default: return <SparkleRegular className={styles.suggestionIcon} />;
    }
  };
  
  return (
    <div className={styles.emptyChatContainer}>
      <div className={styles.header}>
        <Bot24Filled className={styles.botIcon} />
        <Text className={styles.title}>Excel AI Assistant</Text>
        <Text className={styles.subtitle}>
          Ask questions about your data and Excel tasks
        </Text>
      </div>
      
      <div className={styles.suggestionsContainer}>
        {suggestions.slice(0, 3).map((suggestion, index) => (
          <Button 
            key={index}
            className={styles.suggestionButton}
            onClick={() => onSuggestionClick(suggestion)}
          >
            <div className={styles.suggestionContent}>
              {getIcon(index)}
              <span className={styles.suggestionText}>{suggestion}</span>
            </div>
            <ArrowRight16Filled className={styles.arrowIcon} />
          </Button>
        ))}
      </div>
      
      <Text className={styles.footer}>
        Choose a suggestion or type your question
      </Text>
    </div>
  );
};

export default EmptyChat; 