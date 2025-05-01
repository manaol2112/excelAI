  import React from "react";
import {
  makeStyles,
  tokens,
  Button
} from "@fluentui/react-components";

// Styles for the suggestions list component
const useStyles = makeStyles({
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

// Suggestions list component
const SuggestionsList = ({ suggestions, onSuggestionClick }) => {
  const styles = useStyles();
  
  // If no suggestions, don't render anything
  if (!suggestions || suggestions.length === 0) {
    return null;
  }
  
  return (
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
  );
};

export default SuggestionsList; 