import React from "react";
import {
  makeStyles,
  tokens,
  Text,
  Button,
  Divider,
  mergeClasses,
  useId
} from "@fluentui/react-components";
import { 
  Bot24Regular, 
  ArrowForward20Regular, 
  SparkleRegular, 
  DocumentSearchRegular, 
  CalculatorRegular,
  DataTrendingRegular
} from "@fluentui/react-icons";

const useStyles = makeStyles({
  emptyChatMessage: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    height: "100%",
    color: tokens.colorNeutralForeground1,
    padding: "40px 24px",
    textAlign: "center",
    position: "relative",
    overflow: "hidden",
    background: `linear-gradient(145deg, ${tokens.colorNeutralBackground1} 0%, ${tokens.colorNeutralBackground2} 50%, ${tokens.colorNeutralBackground1} 100%)`,
    "&::before": {
      content: '""',
      position: "absolute",
      top: 0,
      left: 0,
      right: 0,
      height: "6px",
      background: "linear-gradient(90deg, #0078d4, #2b88d8, #4a9edf, #6cb4e7)",
      zIndex: 1,
    },
  },
  backgroundPattern: {
    position: "absolute",
    top: 0,
    left: 0,
    right: 0,
    bottom: 0,
    opacity: 0.03,
    backgroundImage: "radial-gradient(circle at 25px 25px, #0078d4 2px, transparent 0)",
    backgroundSize: "50px 50px",
    pointerEvents: "none",
    zIndex: 0,
  },
  contentContainer: {
    position: "relative",
    zIndex: 1,
    width: "100%",
    maxWidth: "800px",
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
  },
  welcomeContainer: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    gap: "16px",
    maxWidth: "600px",
    marginBottom: "32px",
  },
  robotIconContainer: {
    position: "relative",
    marginBottom: "24px",
  },
  robotIcon: {
    fontSize: "72px",
    color: tokens.colorBrandForeground1,
    padding: "20px",
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: "50%",
    boxShadow: `
      0 10px 25px -5px rgba(0, 120, 212, 0.2),
      0 0 10px -5px rgba(0, 0, 0, 0.1)
    `,
    position: "relative",
    zIndex: 2,
  },
  iconGlow: {
    position: "absolute",
    top: "-5%",
    left: "-5%",
    right: "-5%",
    bottom: "-5%",
    borderRadius: "50%",
    background: "radial-gradient(circle, rgba(0, 120, 212, 0.15) 0%, rgba(0, 120, 212, 0) 70%)",
    animation: "pulse 3s infinite ease-in-out",
    zIndex: 1,
  },
  '@keyframes pulse': {
    '0%': { transform: 'scale(1)', opacity: 0.7 },
    '50%': { transform: 'scale(1.05)', opacity: 0.9 },
    '100%': { transform: 'scale(1)', opacity: 0.7 },
  },
  titleText: {
    fontSize: "28px",
    lineHeight: "34px",
    fontWeight: "600",
    marginBottom: "8px",
    background: "linear-gradient(90deg, #0078d4 0%, #106ebe 100%)",
    WebkitBackgroundClip: "text",
    WebkitTextFillColor: "transparent",
    letterSpacing: "-0.3px",
  },
  subtitleText: {
    fontSize: tokens.fontSizeBase300,
    color: tokens.colorNeutralForeground2,
    maxWidth: "500px",
    lineHeight: "1.6",
    fontWeight: tokens.fontWeightRegular,
  },
  suggestionSection: {
    width: "100%",
    maxWidth: "800px",
    display: "flex",
    flexDirection: "column",
    gap: "20px",
    marginTop: "12px",
  },
  suggestionsGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fit, minmax(240px, 1fr))",
    gap: "16px",
    width: "100%",
  },
  suggestionCard: {
    display: "flex",
    flexDirection: "column",
    padding: "20px",
    borderRadius: "12px",
    backgroundColor: tokens.colorNeutralBackground1,
    border: `1px solid ${tokens.colorNeutralStrokeAccessible}`,
    transition: "all 0.25s cubic-bezier(0.25, 0.46, 0.45, 0.94)",
    cursor: "pointer",
    position: "relative",
    overflow: "hidden",
    backdropFilter: "blur(8px)",
    boxShadow: "0 4px 12px rgba(0, 0, 0, 0.05)",
    height: "100%",
    
    ':hover': {
      transform: "translateY(-4px)",
      boxShadow: "0 12px 24px rgba(0, 0, 0, 0.08)",
      borderColor: tokens.colorBrandStroke1,
      background: `linear-gradient(145deg, ${tokens.colorNeutralBackground1} 0%, ${tokens.colorNeutralBackground1Hover} 100%)`,
    },
    ':active': {
      transform: "translateY(-2px)",
      transition: "all 0.1s ease",
    },
    '::before': {
      content: '""',
      position: "absolute",
      top: 0,
      left: 0,
      width: "4px",
      height: "100%",
      background: "linear-gradient(to bottom, #0078d4, #106ebe)",
      opacity: 0,
      transition: "opacity 0.3s ease",
    },
    ':hover::before': {
      opacity: 1,
    }
  },
  suggestionTitle: {
    fontWeight: "600",
    marginBottom: "10px",
    display: "flex",
    alignItems: "center",
    gap: "10px",
    fontSize: tokens.fontSizeBase400,
    color: tokens.colorNeutralForeground1,
    transition: "color 0.2s ease",
  },
  suggestionIconWrapper: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    width: "32px",
    height: "32px",
    borderRadius: "8px",
    background: "linear-gradient(145deg, rgba(0, 120, 212, 0.1), rgba(0, 120, 212, 0.2))",
    marginRight: "4px",
    transition: "transform 0.2s ease",
  },
  suggestionIcon: {
    color: tokens.colorBrandForeground1,
    fontSize: "18px",
  },
  sectionTitle: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
    marginBottom: "16px",
    color: tokens.colorNeutralForeground1,
    display: "flex",
    alignItems: "center",
    gap: "10px",
    position: "relative",
    paddingLeft: "12px",
    
    "::before": {
      content: '""',
      position: "absolute",
      left: 0,
      top: "50%",
      transform: "translateY(-50%)",
      width: "4px",
      height: "18px",
      background: "linear-gradient(to bottom, #0078d4, #106ebe)",
      borderRadius: "2px",
    }
  },
  tryText: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    lineHeight: "1.5",
    flexGrow: 1,
  },
  actionButton: {
    marginTop: "12px",
    alignSelf: "flex-end",
    borderRadius: "8px",
    transition: "all 0.2s ease",
    
    ":hover": {
      background: "rgba(0, 120, 212, 0.08)",
    }
  },
  footerText: {
    color: tokens.colorNeutralForeground3,
    marginTop: "20px",
    fontSize: tokens.fontSizeBase200,
    maxWidth: "600px",
    padding: "10px 20px",
    borderRadius: "6px",
    background: "rgba(0, 0, 0, 0.02)",
    backdropFilter: "blur(4px)",
    border: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  divider: {
    width: "100%",
    margin: "20px 0",
    opacity: 0.6,
  }
});

const EmptyChat = ({ mode, suggestions, onSuggestionClick }) => {
  const styles = useStyles();
  const uniqueId = useId("suggestion");
  
  // Group suggestions by category
  const formattedSuggestions = [
    { text: suggestions[0], icon: <DocumentSearchRegular /> },
    { text: suggestions[1], icon: <CalculatorRegular /> },
    { text: suggestions[2], icon: <DataTrendingRegular /> },
    { text: suggestions[3], icon: <SparkleRegular /> }
  ];
  
  return (
    <div className={styles.emptyChatMessage}>
      <div className={styles.backgroundPattern} />
      
      <div className={styles.contentContainer}>
        <div className={styles.welcomeContainer}>
          <div className={styles.robotIconContainer}>
            <div className={styles.iconGlow} />
            <Bot24Regular className={styles.robotIcon} />
          </div>
          <Text className={styles.titleText}>Excel AI Assistant</Text>
          <Text className={styles.subtitleText}>
            Your intelligent companion for Excel tasks. Ask questions, analyze data,
            and let AI help you work smarter with your spreadsheet data.
          </Text>
        </div>
        
        <div className={styles.suggestionSection}>
          <Text className={styles.sectionTitle}>
            <SparkleRegular /> Quick Actions
          </Text>
          <div className={styles.suggestionsGrid}>
            {formattedSuggestions.map((suggestion, index) => (
              <div
                key={`${uniqueId}-${index}`}
                className={styles.suggestionCard}
                onClick={() => onSuggestionClick(suggestion.text)}
                role="button"
                aria-label={suggestion.text}
              >
                <div className={styles.suggestionTitle}>
                  <div className={styles.suggestionIconWrapper}>
                    {suggestion.icon}
                  </div>
                  {suggestion.text.split(' ').slice(0, 3).join(' ')}...
                </div>
                <Text className={styles.tryText}>{suggestion.text}</Text>
                <Button 
                  appearance="transparent" 
                  icon={<ArrowForward20Regular />} 
                  iconPosition="after"
                  className={styles.actionButton}
                >
                  Try it
                </Button>
              </div>
            ))}
          </div>
        </div>
        
        <Divider className={styles.divider} />
        
        <Text className={styles.footerText}>
          Excel AI Assistant will help you analyze your data, create formulas, format content, and generate insights from your spreadsheets.
        </Text>
      </div>
    </div>
  );
};

export default EmptyChat; 