import * as React from "react";
import { useState } from "react";
import {
  Button,
  Field,
  Textarea,
  Text,
  makeStyles,
  tokens,
  Spinner,
  MessageBar,
  MessageBarBody,
  Card,
  Divider,
  CardHeader,
  Avatar
} from "@fluentui/react-components";
import { 
  Calculator24Regular,
  TextboxAlignMiddle24Regular,
  Code24Regular,
  DocumentCopy24Regular, 
  DocumentEdit24Regular
} from "@fluentui/react-icons";
import { useAI } from "../../context/AIContext";
import excelService from "../../services/excelService";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
    padding: "12px",
    background: tokens.colorNeutralBackground1,
    borderRadius: tokens.borderRadiusMedium,
    height: "100%",
  },
  header: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase500,
    padding: "8px 0 16px 0",
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  formulaCard: {
    marginTop: "20px",
    backgroundColor: tokens.colorBrandBackground2,
    borderLeft: `4px solid ${tokens.colorBrandStroke1}`,
    boxShadow: tokens.shadow4,
  },
  formulaContent: {
    whiteSpace: "pre-wrap",
    padding: "12px",
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
  actionButton: {
    marginTop: "6px",
    display: "flex",
    alignItems: "center",
    gap: "6px",
  },
  buttonsContainer: {
    display: "flex",
    gap: "12px",
    marginTop: "20px",
    flexWrap: "wrap",
  },
  errorMessage: {
    marginBottom: "16px",
  },
  explanation: {
    fontSize: tokens.fontSizeBase300,
    color: tokens.colorNeutralForeground1,
    marginTop: "12px",
    lineHeight: tokens.lineHeightBase300,
  },
  inputContainer: {
    marginTop: "16px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    padding: "16px",
    boxShadow: tokens.shadow2,
  },
  generateButton: {
    marginTop: "16px",
    display: "flex",
    justifyContent: "center",
  },
  loadingSpinner: {
    marginRight: "8px",
  },
  textareaField: {
    '& textarea': {
      height: '100px',
      fontFamily: tokens.fontFamilyBase,
      fontSize: tokens.fontSizeBase300,
      padding: "10px 14px",
    }
  },
  cardHeader: {
    paddingBottom: "0",
  },
  empty: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    textAlign: "center",
    padding: "24px",
    color: tokens.colorNeutralForeground2,
    gap: "12px",
  },
  formulaIcon: {
    fontSize: "48px",
    marginBottom: "16px",
    color: tokens.colorBrandForeground1,
  }
});

const FormulaAssistant = () => {
  const [description, setDescription] = useState("");
  const [formula, setFormula] = useState("");
  const [isApplying, setIsApplying] = useState(false);
  const [isCopied, setIsCopied] = useState(false);
  const styles = useStyles();
  const { suggestFormula, isLoading, error, isApiKeyValid } = useAI();

  const handleGenerateFormula = async () => {
    if (!description.trim() || !isApiKeyValid) return;

    try {
      const result = await suggestFormula(description);
      setFormula(result);
      setIsCopied(false);
    } catch (error) {
      console.error("Error generating formula:", error);
    }
  };

  const handleApplyFormula = async () => {
    if (!formula) return;

    setIsApplying(true);
    try {
      // Try to extract the Excel formula from the AI response
      const extractedFormula = extractFormulaCode(formula);
      if (extractedFormula && extractedFormula.startsWith('=')) {
        await excelService.insertFormula(extractedFormula.substring(1), "A1");
      } else if (extractedFormula) {
        await excelService.insertFormula(extractedFormula, "A1");
      } else {
        // If no clear formula pattern, just use the first line with an =
        const lines = formula.split('\n');
        for (const line of lines) {
          if (line.trim().startsWith('=')) {
            const formulaText = line.trim().substring(1); // Remove the = sign
            await excelService.insertFormula(formulaText, "A1");
            break;
          }
        }
      }
    } catch (error) {
      console.error("Error applying formula:", error);
    } finally {
      setIsApplying(false);
    }
  };

  const copyFormulaToClipboard = () => {
    const extractedFormula = extractFormulaCode(formula);
    if (extractedFormula) {
      navigator.clipboard.writeText(extractedFormula);
      setIsCopied(true);
      setTimeout(() => setIsCopied(false), 2000);
    }
  };

  const extractFormulaCode = (content) => {
    if (!content) return null;
    
    // Try to find a code block
    const codeBlockMatch = content.match(/```(?:excel)?\s*([\s\S]*?)```/);
    if (codeBlockMatch && codeBlockMatch[1]) {
      return codeBlockMatch[1].trim();
    }
    
    // Look for lines starting with =
    const lines = content.split('\n');
    for (const line of lines) {
      if (line.trim().startsWith('=')) {
        return line.trim();
      }
    }
    
    return null;
  };

  const renderFormulaContent = () => {
    if (!formula) return null;
    
    const formulaCode = extractFormulaCode(formula);
    
    return (
      <Card className={styles.formulaCard}>
        <CardHeader 
          className={styles.cardHeader}
          image={
            <Avatar 
              icon={<Code24Regular />} 
              color="brand"
            />
          }
          header={
            <Text weight="semibold">Suggested Formula</Text>
          }
        />
        <div className={styles.formulaContent}>
          {formulaCode && (
            <div className={styles.codeBlock}>
              <Text>{formulaCode}</Text>
            </div>
          )}
          
          <Divider style={{ margin: '16px 0' }} />
          
          <Text className={styles.explanation}>{formula}</Text>
          
          <div className={styles.buttonsContainer}>
            <Button 
              appearance="primary" 
              onClick={handleApplyFormula} 
              disabled={isApplying || !formulaCode}
              icon={isApplying ? <Spinner size="tiny" /> : <TextboxAlignMiddle24Regular />}
            >
              Apply to Cell A1
            </Button>
            
            <Button 
              appearance="outline" 
              onClick={copyFormulaToClipboard} 
              disabled={!formulaCode}
              icon={<DocumentCopy24Regular />}
            >
              {isCopied ? "Copied!" : "Copy Formula"}
            </Button>
          </div>
        </div>
      </Card>
    );
  };

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <Calculator24Regular />
        <Text>Formula Assistant</Text>
      </div>

      {error && (
        <MessageBar intent="error" className={styles.errorMessage}>
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}

      {!isApiKeyValid && (
        <MessageBar intent="warning">
          <MessageBarBody>Please set your OpenAI API key to use the formula assistant.</MessageBarBody>
        </MessageBar>
      )}

      {!formula ? (
        <>
          <div className={styles.inputContainer}>
            <Field 
              label="Describe what you want to calculate" 
              hint="E.g., 'Sum values in column B if column A contains 'Sales''"
              className={styles.textareaField}
            >
              <Textarea
                placeholder="Describe the formula you need in natural language..."
                value={description}
                onChange={(e) => setDescription(e.target.value)}
                resize="vertical"
                disabled={!isApiKeyValid}
              />
            </Field>

            <div className={styles.generateButton}>
              <Button 
                appearance="primary" 
                onClick={handleGenerateFormula} 
                disabled={!description.trim() || isLoading || !isApiKeyValid}
                icon={isLoading ? <Spinner className={styles.loadingSpinner} size="tiny" /> : <DocumentEdit24Regular />}
                size="large"
              >
                {isLoading ? "Generating..." : "Generate Formula"}
              </Button>
            </div>
          </div>
          
          {!isLoading && (
            <div className={styles.empty}>
              <Calculator24Regular className={styles.formulaIcon} />
              <Text size={400} weight="semibold">Generate Excel Formulas with AI</Text>
              <Text size={300}>Describe what you want to calculate in plain English and AI will create the formula for you.</Text>
              <Text size={200} style={{ marginTop: '12px' }}>Examples:</Text>
              <Text size={200}>• "Count the number of cells that contain 'Completed'"</Text>
              <Text size={200}>• "Calculate the weighted average of values in column C based on weights in column D"</Text>
              <Text size={200}>• "Lookup a value in a table and return the corresponding result"</Text>
            </div>
          )}
        </>
      ) : (
        renderFormulaContent()
      )}
    </div>
  );
};

export default FormulaAssistant; 