import React from "react";
import {
  makeStyles,
  tokens,
  Text,
  Button,
  Divider
} from "@fluentui/react-components";
import { Bot24Regular } from "@fluentui/react-icons";

const useStyles = makeStyles({
  emptyChatMessage: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    height: "100%",
    gap: "16px",
    color: tokens.colorNeutralForeground2,
    padding: "24px",
    textAlign: "center",
  },
  robotIcon: {
    fontSize: "48px",
    marginBottom: "16px",
    color: tokens.colorBrandForeground1,
  },
  suggestions: {
    display: "flex",
    gap: "8px",
    flexWrap: "wrap",
    marginTop: "20px",
  },
  suggestionButton: {
    display: "flex",
    alignItems: "center",
    gap: "6px",
    padding: "8px 12px",
    borderRadius: tokens.borderRadiusMedium,
    boxShadow: tokens.shadow2,
    transition: "all 0.2s ease",
    ':hover': {
      transform: "translateY(-2px)",
      boxShadow: tokens.shadow8,
    }
  }
});

const EmptyChat = ({ suggestions, onSuggestionClick }) => {
  const styles = useStyles();
  
  return (
    <div className={styles.emptyChatMessage}>
      <Bot24Regular className={styles.robotIcon} />
      <Text size={500} weight="semibold">How can I help with your Excel tasks?</Text>
      <Text size={300}>Ask me about formulas, data analysis, or Excel functions</Text>
      
      <div className={styles.suggestions}>
        {suggestions.map((suggestion, index) => (
          <Button
            key={index}
            appearance="subtle"
            size="medium"
            icon={suggestion.icon}
            className={styles.suggestionButton}
            onClick={() => onSuggestionClick(suggestion)}
          >
            {suggestion}
          </Button>
        ))}
      </div>
      
      <Divider style={{ width: '80%', margin: '20px 0' }} />
      
      <Text size={300} weight="semibold">
        Currently in Agent Mode
      </Text>
      <Text size={200}>
        I can directly apply changes to your spreadsheet when you ask.
      </Text>
      <Text size={200} style={{ marginTop: '8px' }}>
        Try asking me to "color cell A3 yellow" or "apply bold to cell B5".
      </Text>
      <Text size={200} style={{ marginTop: '8px' }}>
        Switch between modes using the toggle in the top-right corner.
      </Text>
    </div>
  );
};

export default EmptyChat; 