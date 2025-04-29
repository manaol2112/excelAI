import * as React from "react";
import {
  Dropdown,
  Option,
  makeStyles,
  tokens,
  Text,
  Field,
  Card,
  CardHeader,
  Avatar,
  Badge
} from "@fluentui/react-components";
import { BookInformation24Regular, CheckmarkCircle24Regular, Star24Regular } from "@fluentui/react-icons";
import { useAI } from "../../context/AIContext";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    padding: "16px",
    background: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    boxShadow: tokens.shadow4,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  header: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase400,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    paddingBottom: "12px",
    marginBottom: "4px",
  },
  modelCards: {
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    marginTop: "12px",
  },
  modelCard: {
    display: "flex",
    alignItems: "center",
    gap: "12px",
    padding: "12px",
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground3,
    cursor: "pointer",
    border: `1px solid transparent`,
    transition: "all 0.2s ease",
    position: "relative",
    overflow: "hidden",
  },
  selectedModel: {
    borderColor: tokens.colorBrandStroke1,
    backgroundColor: tokens.colorBrandBackground2,
  },
  modelIcon: {
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
    padding: "6px",
    borderRadius: "50%",
  },
  modelInfo: {
    display: "flex",
    flexDirection: "column",
    flexGrow: 1,
  },
  modelName: {
    fontWeight: tokens.fontWeightSemibold,
  },
  modelDescription: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    lineHeight: tokens.lineHeightBase200,
  },
  selectedBadge: {
    position: "absolute",
    top: "8px",
    right: "8px",
    backgroundColor: tokens.colorStatusSuccessBackground1,
    color: tokens.colorStatusSuccessForeground1,
    padding: "2px 8px",
    borderRadius: tokens.borderRadiusCircular,
    fontSize: tokens.fontSizeBase100,
  },
  divider: {
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    margin: "12px 0",
  }
});

// Model info
const MODEL_INFO = {
  'gpt-3.5-turbo': {
    icon: <CheckmarkCircle24Regular />,
    description: "Fast and cost-effective model that's great for most tasks.",
    strengths: "Best for: Simple analysis, basic formulas, and quick responses."
  },
  'gpt-4': {
    icon: <BookInformation24Regular />,
    description: "Advanced reasoning with deeper understanding of complex data.",
    strengths: "Best for: Complex analysis, nuanced insights, and detailed explanations."
  },
  'gpt-4-turbo': {
    icon: <Star24Regular />,
    description: "Balances speed with advanced capabilities for optimal performance.",
    strengths: "Best for: Complex tasks requiring fast response times."
  }
};

const ModelSelector = () => {
  const { selectedModel, setSelectedModel, availableModels, selectedProvider } = useAI();
  const styles = useStyles();

  const handleModelSelect = (model) => {
    setSelectedModel(model);
  };

  return (
    <Card className={styles.container}>
      <CardHeader 
        header={
          <div className={styles.header}>
            <BookInformation24Regular />
            <Text>AI Model Selection</Text>
          </div>
        }
      />

      <Text size={300}>
        Choose the AI model that best fits your needs. More powerful models can handle complex requests but may be slower.
      </Text>

      <div className={styles.modelCards}>
        {availableModels[selectedProvider].map((model) => (
          <div 
            key={model.id} 
            className={`${styles.modelCard} ${selectedModel?.id === model.id ? styles.selectedModel : ''}`}
            onClick={() => handleModelSelect(model)}
          >
            <Avatar 
              icon={MODEL_INFO[model.id]?.icon} 
              color="brand" 
              className={styles.modelIcon}
            />
            
            <div className={styles.modelInfo}>
              <Text className={styles.modelName}>{model.name}</Text>
              <Text className={styles.modelDescription}>
                {MODEL_INFO[model.id]?.description}
              </Text>
            </div>
            
            {selectedModel?.id === model.id && (
              <Badge appearance="filled" className={styles.selectedBadge}>
                Selected
              </Badge>
            )}
          </div>
        ))}
      </div>
    </Card>
  );
};

export default ModelSelector; 