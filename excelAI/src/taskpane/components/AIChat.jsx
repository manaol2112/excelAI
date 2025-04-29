import * as React from "react";
import { useState, useRef, useEffect } from "react";
import {
  Button,
  Input,
  Text,
  makeStyles,
  tokens,
  Spinner,
  MessageBar,
  MessageBarBody,
  Card,
  CardHeader,
  CardPreview,
  Avatar,
  Badge,
  Divider,
  Menu,
  MenuTrigger,
  MenuPopover,
  MenuList,
  MenuItem,
  Tooltip
} from "@fluentui/react-components";
import { 
  Send24Regular, 
  Bot24Regular,
  Person24Regular, 
  Calculator24Regular,
  Lightbulb24Regular,
  ChevronDown20Regular,
  AppsAddIn24Regular,
  TextboxAlignMiddle24Regular,
  DocumentTable24Regular,
  ArrowSortDown24Regular,
  ChevronRight24Regular,
  ArrowRotateClockwise24Regular,
  ArrowRotateClockwise20Regular,
  ArrowUndo24Regular
} from "@fluentui/react-icons";
import { useAI } from "../../context/AIContext";
import excelService from "../../services/excelService";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    height: "100%",
    gap: "16px",
    padding: "0 12px",
  },
  chatHeader: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: "8px",
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase500,
    padding: "12px 0",
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  headerLeft: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
  },
  headerRight: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
  },
  undoButton: {
    display: "flex",
    alignItems: "center",
    gap: "4px",
    fontSize: tokens.fontSizeBase200,
    padding: "4px 8px",
    height: "28px",
    backgroundColor: tokens.colorNeutralBackground2,
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground3,
    }
  },
  chatContainer: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
    overflow: "auto",
    flexGrow: 1,
    padding: "8px 0",
  },
  messageInput: {
    display: "flex",
    gap: "8px",
    alignItems: "flex-start",
    marginTop: "8px",
    marginBottom: "8px",
  },
  textarea: {
    width: "100%",
    resize: "none",
    borderRadius: tokens.borderRadiusMedium,
    padding: "10px 14px",
    fontFamily: tokens.fontFamilyBase,
    fontSize: tokens.fontSizeBase300,
  },
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
  },
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
  loadingContainer: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    padding: "12px",
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground2,
  },
  errorMessage: {
    marginBottom: "16px",
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
  actionButtonsGroup: {
    display: "flex",
    gap: "8px",
    marginTop: "12px",
    flexWrap: "wrap",
  },
  actionButton: {
    marginTop: "6px",
    display: "flex",
    alignItems: "center",
    gap: "6px",
  },
  modeSelector: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    padding: "4px 8px",
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground2,
    cursor: "pointer",
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground3,
    }
  },
  avatar: {
    boxShadow: tokens.shadow4,
  }
});

// Sample suggestions for Excel
const EXCEL_SUGGESTIONS = [
  { text: "Create a summary of selected data", icon: <Calculator24Regular /> },
  { text: "Suggest a formula for this calculation", icon: <Lightbulb24Regular /> },
  { text: "Format my table professionally", icon: <DocumentTable24Regular /> },
  { text: "Generate a chart from selected data", icon: <ChevronRight24Regular /> },
  { text: "Create a table from my data", icon: <DocumentTable24Regular /> },
  { text: "Apply conditional formatting to my data", icon: <DocumentTable24Regular /> },
  { text: "Sort and filter this dataset", icon: <ArrowSortDown24Regular /> },
  { text: "Create a new worksheet", icon: <AppsAddIn24Regular /> },
];

// AI operation modes
const AI_MODES = {
  ASK: { name: "Ask Mode", description: "Get answers in chat without modifying your spreadsheet" },
  AGENT: { name: "Agent Mode", description: "AI will directly apply changes to your spreadsheet" },
  PROMPT: { name: "Prompt Mode", description: "Ask permission before applying changes" }
};

const AIChat = () => {
  const [messages, setMessages] = useState([]);
  const [inputValue, setInputValue] = useState("");
  const [isProcessingAction, setIsProcessingAction] = useState(false);
  const [aiMode, setAiMode] = useState("AGENT"); // Changed default from "ASK" to "AGENT"
  const [isResendMode, setIsResendMode] = useState(false);
  const [conversationContext, setConversationContext] = useState({
    lastMentionedRange: null,
    lastAction: null,
    pendingQuestion: null
  });
  // Add conversation history state to track context over time
  const [conversationHistory, setConversationHistory] = useState({
    recentCells: [], // Recently mentioned cells/ranges
    recentActions: [], // Recently performed actions
    worksheetState: null, // Latest worksheet state information
    lastActiveSheet: null, // Last active worksheet
  });
  const [isUndoAvailable, setIsUndoAvailable] = useState(false);
  const [isUndoInProgress, setIsUndoInProgress] = useState(false);
  const chatContainerRef = useRef(null);
  const styles = useStyles();
  const { generateText, isLoading, error, isApiKeyValid } = useAI();
  
  // Check for available operations in the history to determine if undo is available
  useEffect(() => {
    const checkUndoAvailability = async () => {
      try {
        // Get operations history
        const operationsHistory = excelService.getOperationsHistory();
        console.log("Checking undo availability:", operationsHistory);
        setIsUndoAvailable(operationsHistory && operationsHistory.length > 0);
      } catch (error) {
        console.error("Error checking undo availability:", error);
        setIsUndoAvailable(false);
      }
    };
    
    // Check for undo availability on load and whenever an action is performed
    checkUndoAvailability();
    
    // Also check periodically in case of external changes
    const interval = setInterval(checkUndoAvailability, 5000);
    
    return () => clearInterval(interval);
  }, [isProcessingAction, messages.length]); // Added messages.length as a dependency

  useEffect(() => {
    // Scroll to bottom of chat when new messages are added
    if (chatContainerRef.current) {
      chatContainerRef.current.scrollTop = chatContainerRef.current.scrollHeight;
    }
  }, [messages]);

  // Handle undo operation
  const handleUndo = async () => {
    if (!isUndoAvailable || isUndoInProgress) return;
    
    setIsUndoInProgress(true);
    try {
      const result = await excelService.undoLastOperation();
      
      if (result.success) {
        // Notify user of successful undo
        handleAddMessage({
          role: 'assistant',
          content: `I've undone the last action: ${result.details.type}`,
          isSuccess: true
        });
        
        // Update undo availability
        const operationsHistory = excelService.getOperationsHistory();
        setIsUndoAvailable(operationsHistory && operationsHistory.length > 0);
        
        // Update conversation history to reflect the undo
        updateConversationHistory({
          recentActions: [`Undid ${result.details.type} operation`]
        });
      } else {
        // Handle undo failure
        handleAddMessage({
          role: 'assistant',
          content: `Failed to undo: ${result.message}`,
          isError: true
        });
      }
    } catch (error) {
      console.error("Error during undo:", error);
      handleAddMessage({
        role: 'assistant',
        content: `Error during undo: ${error.message}`,
        isError: true
      });
    } finally {
      setIsUndoInProgress(false);
    }
  };

  // Add this function to test undo functionality
  const handleTestUndo = async () => {
    try {
      const result = await excelService.addTestOperation();
      console.log("Test operation result:", result);
      
      if (result.success) {
        // Check operations history after adding test operation
        const operationsHistory = excelService.getOperationsHistory();
        console.log("Operations history after test:", operationsHistory);
        setIsUndoAvailable(operationsHistory && operationsHistory.length > 0);
        
        // Show confirmation message
        handleAddMessage({
          role: 'assistant',
          content: "Added test operation to undo history. You can now test the Undo button.",
          isSuccess: true
        });
      }
    } catch (error) {
      console.error("Error adding test operation:", error);
    }
  };

  const handleSendMessage = async () => {
    if (!inputValue.trim() || !isApiKeyValid) return;

    // Check if the input is just a cell range without further instructions
    const isCellRangeOnly = /^[A-Z]+[0-9]+(-|:)[A-Z]+[0-9]+$/i.test(inputValue.trim());
    
    // If this is just a cell range and we have a pending question, combine them
    let effectiveInput = inputValue;
    if (isCellRangeOnly && conversationContext.pendingQuestion) {
      effectiveInput = `${conversationContext.pendingQuestion} ${inputValue}`;
      console.log(`Using pending question with provided range: ${effectiveInput}`);
    }
    
    // Check if user is requesting a calculation - define this early to avoid reference errors
    const isCalculationRequest = 
      /what('s| is)? (the )?(sum|average|mean|total|count|median|min|max)/i.test(effectiveInput) ||
      /(calculate|compute|find|tell me|give me) (the )?(sum|average|mean|total|count|median|min|max)/i.test(effectiveInput) ||
      /=\s*sum\(/i.test(effectiveInput) ||
      /(add|sum) (up|these|this|the values|the numbers)/i.test(effectiveInput);
    
    const userMessage = {
      id: Date.now(),
      role: "user",
      content: inputValue, // Keep the original input in the message display
    };

    setMessages((prev) => [...prev, userMessage]);
    setInputValue("");

    try {
      // Get comprehensive worksheet context FIRST to understand the environment
      let worksheetContext = null;
      try {
        worksheetContext = await excelService.getWorksheetContext();
        console.log("Retrieved worksheet context:", worksheetContext);
        
        // Update conversation history with worksheet information
        if (worksheetContext && worksheetContext.success) {
          updateConversationHistory({
            worksheetState: {
              usedRange: worksheetContext.context.activeSheet.usedRange,
              rowCount: worksheetContext.context.activeSheet.usedRange.rowCount,
              columnCount: worksheetContext.context.activeSheet.usedRange.columnCount
            },
            lastActiveSheet: worksheetContext.context.activeSheet.name
          });
        }
      } catch (contextError) {
        console.error("Error retrieving worksheet context:", contextError);
      }
      
      // Check if there's a selected range or cell specification in the message
      const extractedCellRef = extractCellReference(effectiveInput);
      let selectedRange = null;
      let noSelectionDetected = true;
      
      // First update history if cell reference was found in the message
      if (extractedCellRef) {
        noSelectionDetected = false;
        updateConversationHistory({
          recentCells: [extractedCellRef]
        });
        
        setConversationContext(prev => ({
          ...prev,
          lastMentionedRange: extractedCellRef
        }));
      }
      
      // Next, check if there's an active selection in Excel
      if (worksheetContext?.success && worksheetContext.context.selection && worksheetContext.context.selection.address) {
        noSelectionDetected = false;
        selectedRange = {
          address: worksheetContext.context.selection.address,
          values: worksheetContext.context.selection.values,
          rowCount: worksheetContext.context.selection.rowCount,
          columnCount: worksheetContext.context.selection.columnCount
        };
        
        // Update history with current selection
        updateConversationHistory({
          recentCells: [selectedRange.address]
        });
      } else if (noSelectionDetected) {
        // Attempt to get the current selection directly
        try {
          selectedRange = await excelService.getSelectedRange();
          if (selectedRange && selectedRange.address) {
            noSelectionDetected = false;
            updateConversationHistory({
              recentCells: [selectedRange.address]
            });
          }
        } catch (error) {
          console.log("No selected range detected:", error);
        }
      }
      
      // If still no selection and we need one, check conversation history
      if (noSelectionDetected && conversationHistory.recentCells.length > 0) {
        // Use the most recent cell from history
        const mostRecentCell = conversationHistory.recentCells[0];
        console.log(`No current selection, using most recent from history: ${mostRecentCell}`);
        
        // Check if the cell is still valid in the current sheet
        try {
          // Request formulas specifically when checking historical cell validity
          const rangeInfo = await excelService.getData(mostRecentCell); // Assuming getData returns values/formulas if available
          if (rangeInfo && rangeInfo.values) {
            selectedRange = {
              address: mostRecentCell,
              values: rangeInfo.values,
              rowCount: rangeInfo.values.length,
              columnCount: rangeInfo.values[0].length,
              formulas: rangeInfo.formulas // Add formulas if returned by getData
            };
            noSelectionDetected = false;
          }
        } catch (error) {
          console.error("Error getting worksheet data:", error);
        }
      }
      
      // **Check for Formula Update Request**
      const isFormulaUpdateRequest = 
        (/update formula|modify formula|add condition to formula|change formula/i.test(effectiveInput)) &&
        (extractedCellRef || selectedRange);
        
      let existingFormulaInfo = null;
      if (isFormulaUpdateRequest) {
        const targetRangeAddress = extractedCellRef || selectedRange?.address;
        if (targetRangeAddress) {
          try {
            console.log(`Detected formula update request for range: ${targetRangeAddress}. Fetching existing formula.`);
            // Fetch the range info, specifically asking for formulas
            const rangeData = await Excel.run(async (context) => {
              const sheet = context.workbook.worksheets.getActiveWorksheet();
              const range = sheet.getRange(targetRangeAddress);
              range.load("formulas, rowCount, columnCount");
              await context.sync();
              return { 
                formulas: range.formulas,
                rowCount: range.rowCount,
                columnCount: range.columnCount
              };
            });

            if (rangeData && rangeData.formulas) {
              // Format the existing formula info for the prompt
              // Handle single vs multiple cells
              if (rangeData.rowCount === 1 && rangeData.columnCount === 1) {
                existingFormulaInfo = `Existing formula in ${targetRangeAddress}: ${rangeData.formulas[0][0]}`;
              } else {
                // For multi-cell ranges, maybe just provide the top-left cell formula or a summary
                existingFormulaInfo = `Existing formula in top-left cell (${targetRangeAddress}): ${rangeData.formulas[0][0]}. Note: applies to a range.`; 
                // Or potentially: JSON.stringify(rangeData.formulas)
              }
              console.log("Fetched existing formula info:", existingFormulaInfo);
            }
          } catch (formulaError) {
            console.error(`Error fetching existing formula for ${targetRangeAddress}:`, formulaError);
            existingFormulaInfo = `Could not fetch existing formula for ${targetRangeAddress}.`;
          }
        }
      }
      
      // If still no range detected, try to get the used range of the sheet
      let usedRangeAddress = null;
      if (noSelectionDetected) {
        try {
          const usedRangeResult = await excelService.getUsedRange();
          if (usedRangeResult.success && !usedRangeResult.isEmpty) {
            usedRangeAddress = usedRangeResult.address;
            console.log(`No specific range provided, automatically using used range: ${usedRangeAddress}`);
            // Update conversation context with the detected used range
            setConversationContext(prev => ({
              ...prev,
              lastMentionedRange: usedRangeAddress 
            }));
            updateConversationHistory({
              recentCells: [usedRangeAddress]
            });
            // We treat this as the effective selection for the prompt context
            selectedRange = { address: usedRangeAddress }; 
            noSelectionDetected = false; // A range has been identified
          } else {
            console.log("Worksheet appears to be empty or used range couldn't be determined.");
          }
        } catch (error) {
          console.error("Error getting used range:", error);
        }
      }
      
      // If no selection could be determined, we'll analyze the whole worksheet
      let allData = null;
      
      // Detect if query contains a data lookup request (e.g., "How many X does Y have?")
      const containsDataLookupRequest = /how (many|much)|quantity|total|sum|value|amount|number of/i.test(effectiveInput);
      const containsNameOrEntity = /\b[A-Z][a-z]+ [A-Z][a-z]+\b|\b[A-Z][a-z]+\b/g.test(effectiveInput);
      const isDataAnalysisRequest = containsDataLookupRequest && containsNameOrEntity;
      
      if (noSelectionDetected || 
          isDataAnalysisRequest ||
          effectiveInput.toLowerCase().includes("all cells") || 
          effectiveInput.toLowerCase().includes("all data") ||
          effectiveInput.toLowerCase().includes("entire sheet") ||
          effectiveInput.toLowerCase().includes("entire worksheet")) {
        try {
          allData = await excelService.getAllData();
          console.log("Getting all worksheet data due to no specific selection or data analysis request");
        } catch (error) {
          console.error("Error getting worksheet data:", error);
        }
      }
      
      // Start building the user prompt
      let userPrompt = effectiveInput;
      let calculationResult = null;
      
      // If we have worksheet context, include basic info in the prompt
      if (worksheetContext && worksheetContext.success) {
        const ctx = worksheetContext.context;
        userPrompt += `\n\nCurrent Excel context:
- Workbook: ${ctx.workbook.name}
- Active sheet: ${ctx.activeSheet.name} (${ctx.worksheets.count} total sheets)
- Used range: ${ctx.activeSheet.usedRange.address} (${ctx.activeSheet.usedRange.rowCount} rows x ${ctx.activeSheet.usedRange.columnCount} columns)
- Current selection: ${ctx.selection.address || "None"}
- Tables: ${ctx.tables.count}, Charts: ${ctx.charts.count}`;

        // For debug requests, include much more comprehensive context
        if (effectiveInput.toLowerCase().includes("debug") ||
            effectiveInput.toLowerCase().includes("what is the context") ||
            effectiveInput.includes("show context") ||
            effectiveInput.includes("display context")) {
          userPrompt += `\n\nDetailed Excel context for debugging:
Workbook: ${JSON.stringify(ctx.workbook)}
Worksheets: ${JSON.stringify(ctx.worksheets)}
Active Sheet: ${JSON.stringify(ctx.activeSheet)}
Selection: ${JSON.stringify(ctx.selection)}
Tables: ${JSON.stringify(ctx.tables)}
Charts: ${JSON.stringify(ctx.charts)}
Sample Data: ${JSON.stringify(ctx.activeSheet.sampleData)}`;
        }
      }
      
      // For data analysis requests, add specific instructions to the model
      if (isDataAnalysisRequest) {
        userPrompt += `\n\nIMPORTANT: This is a data analysis request. Please follow these instructions:
        
1. Analyze the data to answer the question directly about "${effectiveInput}"
2. DO NOT ask for more information about column positions or names - use the provided data to figure this out
3. Look for patterns, headers, and relationships in the data to identify the relevant information
4. If names of people/entities are mentioned in the query, search for these in the data
5. When you find the answer, provide it clearly and directly
6. If multiple interpretations are possible, provide the most likely answer based on the data structure
7. Explain briefly how you found the answer in the data (what columns/rows were used)
8. If the answer cannot be determined from the data, explain why

Remember to be confident and direct in your response - users want clear, actionable insights.`;
      }
      
      // For data analysis requests, always include either selected data or all worksheet data
      if (isDataAnalysisRequest) {
        if (selectedRange && selectedRange.values) {
          userPrompt += `\n\nSelected data (${selectedRange.address}):\n${JSON.stringify(selectedRange.values)}`;
        } else if (allData && allData.success && !allData.isEmpty) {
          userPrompt += `\n\nWorksheet data (${allData.address}):\n${JSON.stringify(allData.values)}`;
        } else {
          userPrompt += "\n\nI couldn't access any data from the worksheet to analyze.";
        }
      }
      // If user is asking about selected data, include it in the prompt
      else if (isCalculationRequest) {
        try {
          // If we already have selection info from context, use that first
          if (worksheetContext?.success && 
              worksheetContext.context.selection.address && 
              worksheetContext.context.selection.values.length > 0) {
            
            selectedRange = {
              address: worksheetContext.context.selection.address,
              values: worksheetContext.context.selection.values,
              rowCount: worksheetContext.context.selection.rowCount,
              columnCount: worksheetContext.context.selection.columnCount
            };
            
            userPrompt += `\n\nSelected data (${selectedRange.address}):\n${JSON.stringify(selectedRange.values)}`;
          } else {
            // Otherwise fetch it directly
            selectedRange = await excelService.getSelectedRange();
            if (selectedRange && selectedRange.values) {
              userPrompt += `\n\nSelected data (${selectedRange.address}):\n${JSON.stringify(selectedRange.values)}`;
            }
          }
          
          // If it's a calculation request, perform the calculation
          if (isCalculationRequest && aiMode === "AGENT" && selectedRange) {
            // Determine which calculation to perform
            if (/(sum|total|add up)/i.test(effectiveInput)) {
              calculationResult = await excelService.calculateSum();
              userPrompt += `\n\nI have calculated the sum of the selected range (${calculationResult.address}): ${calculationResult.sum}`;
              
              if (calculationResult.hasNonNumericCells) {
                userPrompt += `\nNote: ${calculationResult.nonNumericCells} out of ${calculationResult.cellCount} cells contain non-numeric values that were excluded from the calculation.`;
              }
            } else if (/(average|mean)/i.test(effectiveInput)) {
              calculationResult = await excelService.calculateAverage();
              userPrompt += `\n\nI have calculated the average of the selected range (${calculationResult.address}): ${calculationResult.average}`;
              
              if (calculationResult.hasNonNumericCells) {
                userPrompt += `\nNote: ${calculationResult.nonNumericCells} out of ${calculationResult.cellCount} cells contain non-numeric values that were excluded from the calculation.`;
              }
            } else if (/count/i.test(effectiveInput)) {
              calculationResult = await excelService.countCells();
              userPrompt += `\n\nI have counted the cells in the selected range (${calculationResult.address}):
              Total cells: ${calculationResult.totalCells}
              Numeric cells: ${calculationResult.numberCells}
              Text cells: ${calculationResult.textCells}
              Non-empty cells: ${calculationResult.nonEmptyCells}`;
            } else {
              // For other statistics (median, min, max, etc.)
              calculationResult = await excelService.getStatistics();
              userPrompt += `\n\nI have calculated statistics for the selected range (${calculationResult.address}):
              Sum: ${calculationResult.sum}
              Average: ${calculationResult.average}
              Median: ${calculationResult.median}
              Min: ${calculationResult.min !== null ? calculationResult.min : 'N/A'}
              Max: ${calculationResult.max !== null ? calculationResult.max : 'N/A'}
              Count: ${calculationResult.count}`;
            }
          }
        } catch (error) {
          console.error("Error fetching selected range:", error);
          userPrompt += "\n\nI tried to access the selected range but encountered an error. Please ensure you have selected a valid range of cells.";
        }
      }
      
      // If user is asking about all data or for searching, get all data from the worksheet
      else if ((effectiveInput.toLowerCase().includes("all cells") || 
          effectiveInput.toLowerCase().includes("all data") ||
          effectiveInput.toLowerCase().includes("entire sheet") ||
          effectiveInput.toLowerCase().includes("entire worksheet")) && !effectiveInput.toLowerCase().includes("debug")) {
        try {
          console.log("Getting all data from worksheet");
          
          // If we already have this from context, use it
          if (worksheetContext?.success && 
              worksheetContext.context.activeSheet.usedRange.address) {
            
            // We already have sample data in context, but for a full response we need to get all data
            allData = await excelService.getAllData();
          } else {
            // Otherwise fetch it directly
            allData = await excelService.getAllData();
          }
          
          if (allData && allData.success && !allData.isEmpty) {
            userPrompt += `\n\nAll data from the worksheet (${allData.address}):\n`;
            
            // Limit the data included in the prompt if it's too large
            if (allData.values.length * allData.values[0].length > 100) {
              userPrompt += `[Data includes ${allData.rowCount} rows and ${allData.columnCount} columns - showing a preview]\n`;
              // Include a preview of the data (first 5 rows and columns)
              const previewRows = Math.min(5, allData.values.length);
              const previewCols = Math.min(5, allData.values[0].length);
              
              const preview = [];
              for (let i = 0; i < previewRows; i++) {
                const row = [];
                for (let j = 0; j < previewCols; j++) {
                  row.push(allData.values[i][j]);
                }
                preview.push(row);
              }
              
              userPrompt += JSON.stringify(preview);
              userPrompt += "\n[Additional data not shown]";
            } else {
              userPrompt += JSON.stringify(allData.values);
            }
          } else {
            userPrompt += "\n\nThe worksheet appears to be empty or I couldn't access the data.";
          }
        } catch (error) {
          console.error("Error getting all data:", error);
          userPrompt += "\n\nI encountered an error while trying to get all data from the worksheet. Please ensure the worksheet is accessible.";
        }
      }
      
      // Add information about the current mode to the prompt
      userPrompt += `\n\nCurrent AI Mode: ${aiMode}`;
      
      // Check for follow-up actions without explicit range
      const isFollowUpAction = /^(apply|set|make|color|format|bold|italic|underline)/i.test(effectiveInput.trim()) && !extractedCellRef;
      const lastKnownRange = conversationHistory.recentCells.length > 0 ? conversationHistory.recentCells[0] : null;

      if (aiMode === "AGENT" && isFollowUpAction && lastKnownRange) {
        userPrompt += `\n\nIMPORTANT: This looks like a follow-up command. Apply the requested action to the last mentioned range: ${lastKnownRange}`;
        console.log(`Detected follow-up action, adding context for AI to use last range: ${lastKnownRange}`);
      } else if (conversationContext.lastMentionedRange && !isCellRangeOnly) {
        userPrompt += `\n\nPreviously mentioned cell range: ${conversationContext.lastMentionedRange}`;
      }
      
      // **Add existing formula info to prompt if fetched**
      if (existingFormulaInfo) {
        userPrompt += `\n\n${existingFormulaInfo}`;
      }
      
      // Add the detected used range to the prompt context if it was used
      if (usedRangeAddress && !extractedCellRef && !selectedRange?.address) {
        userPrompt += `\n\nDetected used range (since no specific range was provided): ${usedRangeAddress}`;
      }
      
      // Add historical context to the prompt
      userPrompt += getHistoricalContext();
      
      if (aiMode === "AGENT") {
        // Much clearer AGENT mode instructions
        userPrompt += `\n\nIMPORTANT: You are in AGENT mode which means you can directly execute supported Excel operations using specific function calls. You should intelligently decide whether to:
        
1. PROVIDE INFORMATION - When the user asks for data, calculations, or explanations
2. PERFORM ACTIONS - When the user wants to modify the spreadsheet using a supported function

Supported Actions:
- Format a range professionally (headers, borders, alignment, autofit): Use keywords like "format table professionally".
- Apply basic formatting (color, bold, italic, underline, alignment, borders): Use keywords like "set background color to blue", "apply bold", "center align".
- Apply conditional formatting based on cell content: For example, "color cells containing 'Electronics' in column E yellow".
- Create tables from data: Provide data in a code block.
- Insert formulas: Specify the formula and target cell.
- Create charts: Specify the data range and optionally the chart type.
- Create or rename worksheets.

For information requests (like "what is the sum of A1:A5?"), respond with the calculated value and a brief explanation.

When suggesting an action, clearly state what you will do and the target range. For example:
"I'll format the table in range A1:L11 professionally."
"I will apply bold formatting to cells B2:B10."
"I will color cells containing 'Electronics' in column E yellow."
"I can create a column chart using the data in A1:B5."

If the user doesn't specify a cell or range, use the currently selected range or the automatically detected used range.

For conditional formatting or filtering requests, make sure to:
1. Extract the specific condition (e.g., cells containing "Electronics")
2. Use the formatCellsByContent function for conditional formatting based on cell contents
3. Do NOT apply formatting to entire ranges - only apply it to cells that meet the condition

IMPORTANT: Do NOT attempt to generate or execute arbitrary VBA code. Stick to the supported actions listed above. If a request cannot be mapped to a supported action, explain that the specific operation is not yet supported.`;
        
        // If we extracted a cell reference, make it explicit
        if (extractedCellRef) {
          userPrompt += `\n\nThe user specifically mentioned cell ${extractedCellRef}, so apply any changes to this cell/range.`;
        }
        
        // If we have a selected range, include it
        if (selectedRange) {
          userPrompt += `\n\nThe user has currently selected ${selectedRange.address}.`;
        } else if (worksheetContext?.success && worksheetContext.context.selection.address) {
          userPrompt += `\n\nThe user has currently selected ${worksheetContext.context.selection.address}.`;
        }
        
        // If we have all data, make a note of that
        if (allData && allData.success && !allData.isEmpty) {
          userPrompt += `\n\nI have provided all data from the worksheet. The data covers range ${allData.address} with ${allData.rowCount} rows and ${allData.columnCount} columns.`;
        }
        
        // If we have a calculation result, tell the model to focus on that
        if (calculationResult) {
          userPrompt += `\n\nFocus your response on the calculation I already performed rather than suggesting a formula. The user wants to know the actual result, not how to calculate it themselves.`;
        }
      } else {
        userPrompt += "\nIn ask mode, you should only provide information and not apply changes to the spreadsheet.";
      }
      
      // For conditional formatting requests in user message, suggest the appropriate code
      if (aiMode === "AGENT" &&
         (inputValue.toLowerCase().includes("color cells") ||
          inputValue.toLowerCase().includes("color all") ||
          inputValue.toLowerCase().includes("highlight cells") ||
          inputValue.toLowerCase().includes("highlight all") ||
          inputValue.toLowerCase().includes("conditional format") ||
          (inputValue.toLowerCase().includes("format") && inputValue.toLowerCase().includes("contain")) ||
          (inputValue.toLowerCase().includes("color") && inputValue.toLowerCase().includes("status")) ||
          (inputValue.toLowerCase().includes("highlight") && inputValue.toLowerCase().includes("rows")) ||
          (inputValue.toLowerCase().includes("highlight") && inputValue.toLowerCase().includes("row")) ||
          (inputValue.toLowerCase().includes("format") && inputValue.toLowerCase().includes("where")) ||
          (inputValue.toLowerCase().includes("missing") || 
           inputValue.toLowerCase().includes("empty") || 
           inputValue.toLowerCase().includes("blank")) ||
          (inputValue.toLowerCase().includes("equals") || 
           inputValue.toLowerCase().includes("equal to") || 
           inputValue.toLowerCase().includes("is exactly") || 
           inputValue.toLowerCase().includes("exactly equal"))
         )) {
        
        // Check if it's a row-based formatting request
        const isRowBasedFormatting = 
          inputValue.toLowerCase().includes("row") || 
          inputValue.toLowerCase().includes("rows") || 
          inputValue.toLowerCase().includes("where") ||
          (inputValue.toLowerCase().includes("with") && inputValue.toLowerCase().includes("status")) ||
          inputValue.toLowerCase().includes("missing") ||
          inputValue.toLowerCase().includes("empty") ||
          inputValue.toLowerCase().includes("blank");
        
        // Check if it's an exact match request
        const isExactMatch = 
          inputValue.toLowerCase().includes("equals") || 
          inputValue.toLowerCase().includes("equal to") || 
          inputValue.toLowerCase().includes("is exactly") || 
          inputValue.toLowerCase().includes("exactly equal") ||
          inputValue.toLowerCase().includes("exactly equals") ||
          inputValue.toLowerCase().includes("equal");
        
        // Extract the condition text if possible
        let conditionText = "";
        let targetColumn = "";
        let formatTarget = "";
        let isMissingValue = false;
        
        // Check for missing/empty value conditions first
        if (inputValue.toLowerCase().includes("missing") ||
            inputValue.toLowerCase().includes("empty") ||
            inputValue.toLowerCase().includes("blank")) {
          
          isMissingValue = true;
          conditionText = "missing";
          
          // Try to identify what is missing
          const missingMatch = inputValue.match(/(missing|empty|blank)\s+(\w+)/i);
          if (missingMatch && missingMatch[2]) {
            // Found what's missing (e.g., "missing Email")
            // Extract the entity that's missing
            const missingEntity = missingMatch[2];
            conditionText = "missing";
            
            // Try to infer which column this might be from the entity name
            if (missingEntity.toLowerCase() === "email") {
              targetColumn = "E"; // Educated guess for email column
            } else if (missingEntity.toLowerCase() === "name") {
              targetColumn = "B"; // Educated guess for name column
            } else if (missingEntity.toLowerCase() === "phone") {
              targetColumn = "F"; // Educated guess for phone column
            } else if (missingEntity.toLowerCase() === "address") {
              targetColumn = "G"; // Educated guess for address column
            } else if (missingEntity.toLowerCase() === "id") {
              targetColumn = "A"; // Educated guess for ID column
            }
          }
        } else {
          // Try to extract the search term (text after "containing", "that are", etc.)
          const containingMatch = inputValue.match(/containing\s+["']?([^"']+)["']?/i);
          const thatAreMatch = inputValue.match(/that\s+(are|has|have)\s+["']?([^"']+)["']?/i);
          const withValueMatch = inputValue.match(/with\s+(value|status)\s+["']?([^"']+)["']?/i);
          const equalsMatch = inputValue.match(/equals?\s+["']?([^"']+)["']?/i) || 
                            inputValue.match(/equal to\s+["']?([^"']+)["']?/i) || 
                            inputValue.match(/exactly\s+["']?([^"']+)["']?/i);
          const whereMatch = inputValue.match(/where\s+\w+\s+(is|contains|equals)\s+["']?([^"']+)["']?/i);
          
          if (containingMatch && containingMatch[1]) {
            conditionText = containingMatch[1].trim();
          } else if (thatAreMatch && thatAreMatch[2]) {
            conditionText = thatAreMatch[2].trim();
          } else if (withValueMatch && withValueMatch[2]) {
            conditionText = withValueMatch[2].trim();
          } else if (equalsMatch && equalsMatch[1]) {
            conditionText = equalsMatch[1].trim();
          } else if (whereMatch && whereMatch[2]) {
            conditionText = whereMatch[2].trim();
          }
        }
        
        // Try to extract the target column
        const columnMatch = inputValue.match(/column\s+([A-Z])/i) || 
                          inputValue.match(/in\s+column\s+([A-Z])/i) ||
                          inputValue.match(/where\s+([A-Z])\s+/i);
                          
        if (columnMatch && columnMatch[1]) {
          targetColumn = columnMatch[1];
        }
        
        // Extract formatTarget (cells, rows, etc.)
        if (inputValue.toLowerCase().includes("rows") || inputValue.toLowerCase().includes("row")) {
          formatTarget = "rows";
        } else if (inputValue.toLowerCase().includes("cells") || inputValue.toLowerCase().includes("cell")) {
          formatTarget = "cells";
        } else if (inputValue.toLowerCase().includes("status")) {
          // If status is mentioned without explicit row/cell, assume row formatting
          formatTarget = "rows";
          // If we found "status" but no column, try to detect status column
          if (!targetColumn) {
            // Try to find reference to status
            const statusColumnMatch = inputValue.match(/status\s+(in|from)\s+column\s+([A-Z])/i);
            if (statusColumnMatch && statusColumnMatch[2]) {
              targetColumn = statusColumnMatch[2];
            } else {
              // Default status column to common choices like "C" or "D" if none specified
              targetColumn = "C";
            }
          }
        } else if (isMissingValue) {
          // For missing value conditions, default to row formatting
          formatTarget = "rows";
        } else if (isRowBasedFormatting) {
          formatTarget = "rows";
        } else {
          formatTarget = "cells";
        }
        
        // Prepare a helpful code example based on the request type
        if (isExactMatch && (isRowBasedFormatting || formatTarget === "rows")) {
          userPrompt += `\n\nThis appears to be a request for conditional formatting based on exact value matching. Please use the formatRowsByExactMatch function to apply formatting to entire rows where a cell exactly equals the specified value.

Example Office.js code for this type of request:
\`\`\`javascript
await Excel.run(async (context) => {
  try {
    // Call the excelService to apply conditional formatting to rows based on exact match
    const result = await excelService.formatRowsByExactMatch(
      "A:Z", // Range to format (all columns to style entire rows)
      ${targetColumn ? `"${targetColumn}"` : `"C"`}, // Column to check for the exact match (e.g., status column)
      ${conditionText ? `"${conditionText}"` : `"Delivered"`}, // Value to exactly match
      { 
        fillColor: "yellow",
        // Optional: fontColor: "red",
        // Optional: bold: true
      }
    );
    console.log("Formatted rows:", result);
  } catch (error) {
    console.error("Error:", error);
  }
});
\`\`\``;
        } else if (isRowBasedFormatting || formatTarget === "rows") {
          userPrompt += `\n\nThis appears to be a row-based conditional formatting request. Please use the formatRowsByCondition function to apply formatting to entire rows based on values in a specific column.

Example Office.js code for this type of request:
\`\`\`javascript
await Excel.run(async (context) => {
  try {
    // Call the excelService to apply conditional formatting to rows
    const result = await excelService.formatRowsByCondition(
      ${targetColumn ? `"A:Z"` : `"A:Z"`}, // Range to format (all columns to style entire rows)
      ${targetColumn ? `"${targetColumn}"` : `"C"`}, // Column to check for the condition (e.g., status column)
      ${isMissingValue ? `"missing"` : conditionText ? `"${conditionText}"` : `"Missing Email"`}, // Condition to check for
      { 
        fillColor: "yellow",
        // Optional: fontColor: "red",
        // Optional: bold: true
      }
    );
    console.log("Formatted rows:", result);
  } catch (error) {
    console.error("Error:", error);
  }
});
\`\`\``;
        } else {
          // Cell-based formatting
          userPrompt += `\n\nThis appears to be a cell-based conditional formatting request. Please use the formatCellsByContent function to apply formatting only to cells that match the specific condition.

Example Office.js code for this type of request:
\`\`\`javascript
await Excel.run(async (context) => {
  try {
    // Call the excelService to apply conditional formatting
    const result = await excelService.formatCellsByContent(
      ${targetColumn ? `"${targetColumn}:${targetColumn}"` : `"A:Z"`}, // Range to check
      ${conditionText ? `"${conditionText}"` : `"Delivered"`}, // Text to match
      { 
        fillColor: "yellow",
        // Optional: fontColor: "red",
        // Optional: bold: true
      }
    );
    console.log("Formatted cells:", result);
  } catch (error) {
    console.error("Error:", error);
  }
});
\`\`\`

// Alternatively, you can use Excel's native conditional formatting API:
\`\`\`javascript
await Excel.run(async (context) => {
  try {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(${targetColumn ? `"${targetColumn}:${targetColumn}"` : `"A:Z"`});
    
    // NOTE: Use conditionalFormats (plural) not conditionalFormat (singular)
    // Use the correct enum value from Excel.ConditionalFormatType
    const conditionalFormat = range.conditionalFormats.add("ContainsText");
    conditionalFormat.textComparison.rule = {
      operator: "Contains",
      text: ${conditionText ? `"${conditionText}"` : `"Delivered"`}
    };
    conditionalFormat.format.fill.color = "yellow";
    
    await context.sync();
    console.log("Applied conditional formatting");
  } catch (error) {
    console.error("Error:", error);
  }
});
\`\`\``;
        }
      }
      
      console.log("Sending prompt to AI:", userPrompt);
      const aiResponse = await generateText(userPrompt);
      console.log("Received AI response:", aiResponse);
      
      // Check if AI is asking for a cell range specification
      const askingForRange = /please (specify|provide|tell me) (the )?(cell|range|cells)/i.test(aiResponse) ||
                              /which (cell|range|cells) (are you|do you)/i.test(aiResponse) ||
                              /could you (please )?(specify|tell me|clarify) which cells/i.test(aiResponse);
                              
      if (askingForRange) {
        // Store the current question as pending
        setConversationContext(prev => ({
          ...prev,
          pendingQuestion: userPrompt
        }));
        console.log("Stored pending question awaiting cell range:", userPrompt);
      } else {
        // Clear any pending question since we got a complete response
        setConversationContext(prev => ({
          ...prev,
          pendingQuestion: null
        }));
      }
      
      const aiMessage = {
        id: Date.now() + 1,
        role: "assistant",
        content: aiResponse,
        selectedData: selectedRange,
        extractedCellRef: extractedCellRef,
        calculationResult: calculationResult,
        allData: allData,
        searchResults: null,
        worksheetContext: worksheetContext?.success ? worksheetContext.context : null
      };
      
      // If in AGENT mode, automatically process any detected actions
      if (aiMode === "AGENT") {
        try {
          // Check if the response is just a code block
          const isJustCodeBlock = aiResponse.trim().match(/^```(js|javascript|vba)\s*[\s\S]*?```\s*$/);
          
          // If just a code block, we'll let handleAddMessage and processAgentActions handle it
          if (!isJustCodeBlock) {
            setMessages((prev) => [...prev, aiMessage]);
          }
          
          await processAgentActions(aiResponse, selectedRange, extractedCellRef, calculationResult, isCalculationRequest);
          
          // If it was just a code block, don't do anything else - the code execution and friendly message are handled elsewhere
          if (isJustCodeBlock) {
            // The message is already being added by processAgentActions or handleAddMessage
            return;
          }
        } catch (actionError) {
          console.error("Error processing agent actions:", actionError);
          // Don't add another error message to the chat, the existing one is enough
        }
      } else {
        // Not in AGENT mode, just add the message as-is
        setMessages((prev) => [...prev, aiMessage]);
      }
      
    } catch (error) {
      console.error("Error sending message:", error);
      const errorMessage = {
        id: Date.now() + 1,
        role: "assistant",
        content: `Sorry, I encountered an error: ${error.message}. Please try again or check if Excel is responding correctly.`,
        isError: true,
      };
      setMessages((prev) => [...prev, errorMessage]);
    }
  };

  // Function to extract cell references from text (e.g., A1, B2:C3, etc.)
  const extractCellReference = (text) => {
    if (!text) return null;
    
    // Common cell reference patterns
    const singleCellPattern = /\b([A-Z]+[0-9]+)\b/g;  // Matches A1, B2, AA10, etc.
    const rangeCellPattern = /\b([A-Z]+[0-9]+[-:]?[A-Z]+[0-9]+)\b/g;  // Matches A1:B2, C3:D4, A1-B2 etc.
    
    // First look for ranges (A1:B2 or A1-B2)
    const rangeMatches = text.match(rangeCellPattern);
    if (rangeMatches && rangeMatches.length > 0) {
      // Convert any range with dash to colon format (A1-B2 to A1:B2)
      let range = rangeMatches[0];
      if (range.includes('-')) {
        range = range.replace('-', ':');
      }
      return range;
    }
    
    // Then look for single cells (A1)
    const singleMatches = text.match(singleCellPattern);
    if (singleMatches && singleMatches.length > 0) {
      return singleMatches[0];
    }
    
    return null;
  };

  // Helper function to process agent actions from the AI response
  const processAgentActions = async (aiResponseText, selectedRange, extractedCellRef, calculationResult, isCalculationRequest) => {
    if (!aiResponseText || aiMode === "ASK") return false;
    
    let actionPerformed = false;
    const actionHistory = []; // Track actions for history
    
    try {
      // Auto-execute Office.js code blocks if present
      const officeJsBlockMatch = aiResponseText.match(/```(?:js|javascript)\s*([\s\S]*?)```/i);
      if (officeJsBlockMatch && officeJsBlockMatch[1]) {
        const jsCode = officeJsBlockMatch[1].trim();
        console.log("Agent auto-executing Office.js code block:", jsCode);
        
        // Only auto-execute if in AGENT mode, not in PROMPT mode
        if (aiMode === "AGENT") {
          // Don't suppress the success message - we'll handle it below
          const result = await runOfficeJsCode(jsCode, true); 
          actionPerformed = true;
          actionHistory.push("Executed Office.js code block");
          
          // Check undo availability after executing code
          try {
            const operationsHistory = excelService.getOperationsHistory();
            console.log("Operations after executing code:", operationsHistory);
            setIsUndoAvailable(operationsHistory && operationsHistory.length > 0);
          } catch (error) {
            console.error("Error checking undo after code execution:", error);
          }
          
          // Extract a human-friendly description of what the code does
          const operationType = extractOperationType(jsCode);
          const friendlyMessage = generateFriendlyMessage(operationType, jsCode);
          
          // Add a human-friendly message to the chat
          handleAddMessage({
            role: 'assistant',
            content: friendlyMessage,
            isSuccess: true
          });
        }
        // In PROMPT mode, execution is handled by the manual button
      }
      
      // Auto-execute VBA code blocks if present
      const vbaBlockMatch = aiResponseText.match(/```vba\s*([\s\S]*?)```/i);
      if (!actionPerformed && vbaBlockMatch && vbaBlockMatch[1]) {
        const vbaCode = vbaBlockMatch[1].trim();
        console.log("Agent auto-executing VBA code block:", vbaCode);
        
        // Only auto-execute if in AGENT mode, not in PROMPT mode
        if (aiMode === "AGENT") {
          const result = await runVbaCode(vbaCode, true); // Pass true to suppress default message
          actionPerformed = true;
          actionHistory.push("Executed VBA code block");
          
          // Add a human-friendly message about VBA execution
          handleAddMessage({
            role: 'assistant',
            content: "I've executed a VBA macro to complete your request.",
            isSuccess: true
          });
        }
        // In PROMPT mode, execution is handled by the manual button
      }

      // REMOVED specific action handlers (create sheet, rename, format range, etc.)
      // The LLM is now responsible for generating code for these actions.

      // Update overall conversation history with action records
      if (actionHistory.length > 0) {
        updateConversationHistory({
          recentActions: actionHistory
        });
      }
    } catch (error) {
      console.error("Error processing agent actions:", error);
      // Add error message to chat if needed, handled in runOfficeJsCode/runVbaCode now
    }
    
    return actionPerformed;
  };

  const parseDataString = (dataString) => {
    // Split by lines and filter out empty lines
    const lines = dataString.split("\n").filter(line => line.trim() !== "");
    // For each line, split by comma or tab and create a row array
    return lines.map(line => line.split(/[,\t]/).map(cell => cell.trim()));
  };

  const pasteDataToExcel = async (data, startCell = "A1") => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(startCell);
        // Get the starting row and column
        const startRow = range.rowIndex;
        const startCol = range.columnIndex;
        
        // Create a range that covers all the data
        const targetRange = sheet.getRangeByIndexes(
          startRow, 
          startCol, 
          data.length, 
          Math.max(...data.map(row => row.length))
        );
        
        // Set the values
        targetRange.values = data;
        
        // Auto-fit columns
        targetRange.format.autofitColumns();
        
        await context.sync();
      });
    } catch (error) {
      console.error("Error pasting data to Excel:", error);
    }
  };

  const handleKeyDown = (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      handleSendMessage();
    }
  };

  const handleSuggestionClick = (suggestion) => {
    setInputValue(suggestion);
    // Focus on the input field
    document.querySelector('input[type="text"]').focus();
  };
  
  // Function to resend a previous message
  const resendMessage = (messageContent) => {
    setInputValue(messageContent);
    setIsResendMode(true);
    // Scroll input into view and focus it
    const inputElement = document.querySelector('input[type="text"]');
    if (inputElement) {
      inputElement.focus();
      inputElement.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
  };

  // Watch for changes to inputValue and clear resend mode if emptied
  useEffect(() => {
    if (inputValue === "") {
      setIsResendMode(false);
    }
  }, [inputValue]);

  const extractCode = (content) => {
    // Simple regex to find Excel formula patterns
    const formulaMatches = content.match(/=\w+\([^)]*\)/g);
    return formulaMatches ? formulaMatches[0] : null;
  };

  const applyFormula = async (content, targetCell = "A1") => {
    setIsProcessingAction(true);
    try {
      const formula = extractCode(content);
      if (formula) {
        // Try to find an explicit cell mention in the content
        const inCellPattern = /in\s+cell\s+([A-Z]+[0-9]+)/i;
        const intoCellPattern = /into\s+cell\s+([A-Z]+[0-9]+)/i; 
        const cellPattern = /cell\s+([A-Z]+[0-9]+)/i;
        
        // Try each pattern
        const inCellMatch = content.match(inCellPattern);
        const intoCellMatch = content.match(intoCellPattern);
        const cellMatch = content.match(cellPattern);
        
        // Use the first match found, or fall back to the provided targetCell
        let formulaTargetCell = targetCell;
        if (inCellMatch && inCellMatch[1]) {
          formulaTargetCell = inCellMatch[1];
        } else if (intoCellMatch && intoCellMatch[1]) {
          formulaTargetCell = intoCellMatch[1];
        } else if (cellMatch && cellMatch[1]) {
          formulaTargetCell = cellMatch[1];
        }
        
        console.log(`Manual formula insertion: ${formula.substring(1)} into cell ${formulaTargetCell}`);
        await excelService.insertFormula(formula.substring(1), formulaTargetCell); // Remove the = sign
      }
    } catch (error) {
      console.error("Error applying formula:", error);
    } finally {
      setIsProcessingAction(false);
    }
  };

  const extractTableData = (content) => {
    const tableMatch = content.match(/```(csv|table)\s*([\s\S]*?)```/i);
    return tableMatch ? tableMatch[2] : null;
  };

  const parseTableData = (tableString) => {
    if (!tableString) return null;
    
    // Split by lines and filter out empty lines
    const rows = tableString.split('\n').filter(line => line.trim());
    
    // Convert each row to an array of cell values
    return rows.map(row => {
      // Handle CSV format (comma-separated)
      if (row.includes(',')) {
        return row.split(',').map(cell => cell.trim());
      }
      // Handle pipe or tab-separated format
      else if (row.includes('|')) {
        // Remove leading/trailing pipes and split
        return row.replace(/^\s*\|\s*|\s*\|\s*$/g, '')
          .split('|')
          .map(cell => cell.trim());
      }
      // Handle tab-separated
      else if (row.includes('\t')) {
        return row.split('\t').map(cell => cell.trim());
      }
      // Handle space-separated as fallback
      else {
        return row.split(/\s{2,}/).map(cell => cell.trim());
      }
    });
  };

  const applyTableData = async (content, targetCell = "A1") => {
    setIsProcessingAction(true);
    try {
      const tableData = extractTableData(content);
      if (tableData) {
        const data = parseTableData(tableData);
        if (data && data.length > 0) {
          await excelService.createTable(data, targetCell);
        }
      }
    } catch (error) {
      console.error("Error applying table data:", error);
    } finally {
      setIsProcessingAction(false);
    }
  };

  // Function to handle professional table formatting
  const formatTableHandler = async (range) => {
    setIsProcessingAction(true);
    try {
      // If range doesn't include a colon (i.e., it's a single cell), try to get the selected range
      if (!range.includes(':')) {
        try {
          const selectedRange = await excelService.getSelectedRange();
          if (selectedRange && selectedRange.success && selectedRange.data.address) {
            range = selectedRange.data.address;
          }
        } catch (error) {
          console.error("Error getting selected range:", error);
        }
      }
      
      // Apply professional formatting with default options
      const options = {
        hasHeaders: true,
        headerFill: "navy",
        headerFont: "white",
        alternateFill: "lightgray",
        autofitColumns: true
      };
      
      await excelService.formatAsTable(range, options);
      
      // Update conversation history
      updateConversationHistory({
        recentCells: [range],
        recentActions: ["Applied professional table formatting to " + range]
      });
      
    } catch (error) {
      console.error("Error formatting table:", error);
    } finally {
      setIsProcessingAction(false);
    }
  };

  const renderMessageContent = (message) => {
    // Check if the message contains code blocks (marked with backticks)
    const codeMatch = message.content.match(/^\s*```(js|javascript|vba)\s*([\s\S]*?)```\s*$/);

    if (codeMatch) {
      const language = codeMatch[1];
      const codeContent = codeMatch[2].trim();
      const isVba = language === "vba";
      const isOfficeJs = language === "js" || language === "javascript";

      // If it's Office.js code, extract operation type and provide a user-friendly message
      if (isOfficeJs && aiMode === "AGENT") {
        const operationType = extractOperationType(codeContent);
        const friendlyMessage = generateFriendlyMessage(operationType, codeContent);
        
        return (
          <>
            <Text className={styles.messageContent}>{friendlyMessage}</Text>
            {/* Include a small indicator that code was executed */}
            <div style={{ 
              marginTop: "8px", 
              fontSize: "12px", 
              color: tokens.colorNeutralForeground3,
              fontStyle: "italic"
            }}>
              <Text>Action performed automatically</Text>
            </div>
          </>
        );
      }

      // If it's just a code block (likely from an action request), don't render text, only the button if VBA
      return (
        <div className={styles.codeBlock}>
          <Text>{codeContent}</Text>
          {isVba && (
            <Button 
              appearance="primary"
              size="small"
              style={{ marginTop: "8px" }}
              onClick={() => runVbaCode(codeContent)}
              disabled={isProcessingAction}
              icon={isProcessingAction ? <Spinner size="tiny" /> : null}
            >
              Execute VBA
            </Button>
          )}
          {isOfficeJs && (
            <Button 
              appearance="primary"
              size="small"
              style={{ marginTop: "8px" }}
              onClick={() => runOfficeJsCode(codeContent)}
              disabled={isProcessingAction}
              icon={isProcessingAction ? <Spinner size="tiny" /> : null}
            >
              Run This Code
            </Button>
          )}
        </div>
      );
    } else if (message.content.includes("```")) {
      // Handle messages with mixed text and code blocks (e.g., explanations with code)
      const parts = message.content.split("```");
      return (
        <>
          {parts.map((part, index) => {
            if (index % 2 === 0) {
              // Regular text
              return <Text key={index} className={styles.messageContent}>{part}</Text>;
            } else {
              // Code block inside text
              const isVba = part.trim().toLowerCase().startsWith("vba");
              const isOfficeJs = part.trim().toLowerCase().startsWith("js") || part.trim().toLowerCase().startsWith("javascript");
              const codeContent = isVba ? part.replace(/^vba\s*\n?/i, "") : isOfficeJs ? part.replace(/^(js|javascript)\s*\n?/i, "") : part;
              return (
                <div key={index} className={styles.codeBlock}>
                  <Text>{codeContent}</Text>
                  {isVba && (
                    <Button 
                      appearance="primary"
                      size="small"
                      style={{ marginTop: "8px" }}
                      onClick={() => runVbaCode(codeContent)}
                      disabled={isProcessingAction}
                      icon={isProcessingAction ? <Spinner size="tiny" /> : null}
                    >
                      Execute VBA
                    </Button>
                  )}
                  {isOfficeJs && !message.content.match(/^\s*```(js|javascript)/) && (
                    <Button 
                      appearance="primary"
                      size="small"
                      style={{ marginTop: "8px", marginLeft: "8px" }}
                      onClick={() => runOfficeJsCode(codeContent)}
                      disabled={isProcessingAction}
                      icon={isProcessingAction ? <Spinner size="tiny" /> : null}
                    >
                      Execute Office.js
                    </Button>
                  )}
                </div>
              );
            }
          })}
          {renderDebugInfo(message)}
        </>
      );
    }

    // Default: Render plain text message (likely analysis/info response)
    return (
      <>
        <Text className={styles.messageContent}>{message.content}</Text>
        {renderDebugInfo(message)}
      </>
    );
  };

  // Helper function to generate a friendly message about what the code does
  const generateFriendlyMessage = (operationType, code) => {
    // Look for more specific information in the code
    let specifics = "";
    
    // Check for filter operations
    if (operationType === "filtering") {
      // Try to extract filter criteria
      const filterMatch = code.match(/apply\(.*?,\s*\d+,\s*["']([^"']+)["']/);
      if (filterMatch && filterMatch[1]) {
        specifics = ` to show only "${filterMatch[1]}" items`;
      } else {
        specifics = " for easier data analysis";
      }
    }
    
    // Check for conditional formatting
    if (operationType === "conditional formatting" || operationType === "cell coloring") {
      // Try to extract the condition and range
      const contentMatch = code.match(/content\s*[:=]\s*["']([^"']+)["']/);
      const rangeMatch = code.match(/range\s*[:=]\s*["']([^"']+)["']/i) || code.match(/getRange\(["']([^"']+)["']\)/);
      const colorMatch = code.match(/fillColor\s*[:=]\s*["']([^"']+)["']/);
      
      if (contentMatch && contentMatch[1]) {
        specifics = ` to cells containing "${contentMatch[1]}"`;
        
        // Add range information if available
        if (rangeMatch && rangeMatch[1]) {
          specifics += ` in range ${rangeMatch[1]}`;
        }
        
        // Add color information if available
        if (colorMatch && colorMatch[1]) {
          specifics += ` with ${colorMatch[1]} color`;
        }
      } else if (code.includes("formatCellsByContent")) {
        // Handle formatCellsByContent function calls
        const rangeContentRegex = /formatCellsByContent\(["']([^"']+)["']\s*,\s*["']([^"']+)["']/;
        const match = code.match(rangeContentRegex);
        
        if (match && match.length >= 3) {
          const range = match[1];
          const content = match[2];
          
          specifics = ` to cells containing "${content}" in range ${range}`;
          
          // Try to extract color information
          if (colorMatch && colorMatch[1]) {
            specifics += ` with ${colorMatch[1]} color`;
          }
        }
      }
    }
    
    // Check for row formatting
    if (operationType === "row formatting") {
      // Try to extract the column reference, condition, and formatting options
      const rowConditionRegex = /formatRowsByCondition\(["']([^"']+)["']\s*,\s*["']([^"']+)["']\s*,\s*["']([^"']+)["']/;
      const match = code.match(rowConditionRegex);
      const colorMatch = code.match(/fillColor\s*[:=]\s*["']([^"']+)["']/);
      
      if (match && match.length >= 4) {
        const range = match[1];
        const columnRef = match[2];
        const condition = match[3];
        
        // Check if it's a missing/empty value condition
        const isMissingCondition = ["missing", "empty", "blank"].includes(condition.toLowerCase());
        
        if (isMissingCondition) {
          specifics = ` to rows with missing values in column ${columnRef}`;
        } else {
          specifics = ` to rows where column ${columnRef} contains "${condition}"`;
        }
        
        // Add color information if available
        if (colorMatch && colorMatch[1]) {
          specifics += ` with ${colorMatch[1]} color`;
        }
      }
    }
    
    // Check for exact match formatting
    if (operationType === "exact match formatting") {
      // Try to extract the column reference, match value, and formatting options
      const exactMatchRegex = /formatRowsByExactMatch\(["']([^"']+)["']\s*,\s*["']([^"']+)["']\s*,\s*["']([^"']+)["']/;
      const match = code.match(exactMatchRegex);
      const colorMatch = code.match(/fillColor\s*[:=]\s*["']([^"']+)["']/);
      
      if (match && match.length >= 4) {
        const range = match[1];
        const columnRef = match[2];
        const exactValue = match[3];
        
        specifics = ` to rows where column ${columnRef} equals "${exactValue}"`;
        
        // Add color information if available
        if (colorMatch && colorMatch[1]) {
          specifics += ` with ${colorMatch[1]} color`;
        }
      }
    }
    
    // Check for specific cell/range references if not already included
    if (!specifics.includes("range")) {
      const rangeMatch = code.match(/getRange\(["']([^"']+)["']\)/);
      if (rangeMatch && rangeMatch[1]) {
        specifics += ` in range ${rangeMatch[1]}`;
      }
    }
    
    // Create the message
    const actions = {
      "filtering": `I've applied a filter${specifics}.`,
      "cell coloring": `I've colored the cells${specifics}.`,
      "conditional formatting": `I've applied conditional formatting${specifics}.`,
      "row formatting": `I've highlighted the rows${specifics}.`,
      "exact match formatting": `I've highlighted rows with exact matches${specifics}.`,
      "text formatting": `I've formatted the text${specifics}.`,
      "chart creation": `I've created a chart${specifics}.`,
      "table creation": `I've created a table${specifics}.`,
      "pivot table creation": `I've created a pivot table${specifics}.`,
      "worksheet creation": `I've created a new worksheet.`,
      "worksheet renaming": `I've renamed the worksheet.`,
      "data entry": `I've entered the data${specifics}.`,
      "formula insertion": `I've inserted the formula${specifics}.`,
      "hiding rows/columns": `I've hidden the rows/columns${specifics}.`,
      "unhiding rows/columns": `I've unhidden the rows/columns${specifics}.`,
    };
    
    return actions[operationType] || `I've completed the ${operationType}${specifics}.`;
  };
  
  // Helper function to extract the operation type from code
  const extractOperationType = (code) => {
    // Default operation type if we can't determine specifics
    let operationType = 'Excel operation';
    
    // Check for our custom formatting functions
    if (/formatCellsByContent/.test(code)) {
      return "conditional formatting";
    }
    
    if (/formatRowsByCondition/.test(code)) {
      return "row formatting";
    }
    
    if (/formatRowsByExactMatch/.test(code)) {
      return "exact match formatting";
    }
    
    // Try to determine what kind of operation was performed
    if (/format\.fill\.color|background|\.fill\s*=/.test(code)) {
      // Check if it's conditional formatting with conditions
      if (/getSpecialCellsOrNullObject|\.values\s*==|\.values\.indexOf|if\s*\(|\.filter\s*\(|contains|includes/.test(code)) {
        operationType = "conditional formatting";
      } else {
        operationType = "cell coloring";
      }
    } else if (/format\.font\.(bold|italic|underline|color)/.test(code)) {
      operationType = "text formatting";
    } else if (/\.merge\(\)/.test(code)) {
      operationType = "cell merging";
    } else if (/\.unmerge\(\)/.test(code)) {
      operationType = "cell unmerging";
    } else if (/format\.(row|column)(Height|Width)\s*=\s*0/.test(code)) {
      operationType = "hiding rows/columns";
    } else if (/format\.(row|column)(Height|Width)\s*=\s*[1-9]/.test(code)) {
      operationType = "unhiding rows/columns";
    } else if (/charts\.add/.test(code)) {
      operationType = "chart creation";
    } else if (/tables\.add/.test(code)) {
      operationType = "table creation";
    } else if (/pivotTables\.add/.test(code)) {
      operationType = "pivot table creation";
    } else if (/\.values\s*=/.test(code)) {
      operationType = "data entry";
    } else if (/\.formulas\s*=/.test(code)) {
      operationType = "formula insertion";
    } else if (/conditionalFormats\.add/.test(code)) {
      operationType = "conditional formatting";
    } else if (/autoFilter\.apply/.test(code)) {
      operationType = "filtering";
    } else if (/\.sort\.apply/.test(code)) {
      operationType = "sorting";
    } else if (/worksheets\.add/.test(code)) {
      operationType = "worksheet creation";
    } else if (/sheet\.name\s*=/.test(code)) {
      operationType = "worksheet renaming";
    } else if (/protection\.protect/.test(code)) {
      operationType = "worksheet protection";
    } else if (/protection\.unprotect/.test(code)) {
      operationType = "worksheet unprotection";
    } else if (/dataValidation\.rule/.test(code)) {
      operationType = "data validation";
    } else if (/hyperlink\s*=/.test(code)) {
      operationType = "hyperlink insertion";
    } else if (/comments\.add/.test(code)) {
      operationType = "comment addition";
    } else if (/range.select\(\)/.test(code)) {
      operationType = "range selection";
    } else if (/slicers\.add/.test(code)) {
      operationType = "slicer creation";
    } else if (/format\.autofitColumns/.test(code)) {
      operationType = "column autofit";
    } else if (/format\.autofitRows/.test(code)) {
      operationType = "row autofit";
    }
    
    return operationType;
  };

  const runVbaCode = async (vbaCode, suppressSuccess = false) => {
    setIsProcessingAction(true);
    try {
      console.log("Executing VBA code:", vbaCode);
      const result = await excelService.runMacro(vbaCode);
      
      // Add human-friendly message if not present
      if (result.success && !result.humanMessage) {
        result.humanMessage = "VBA macro executed successfully.";
      } else if (!result.success && !result.humanMessage) {
        result.humanMessage = `Failed to execute VBA macro: ${result.error || "Unknown error"}`;
      }
      
      // Only add a message to the chat if this is called directly (not through processAgentActions)
      if (!suppressSuccess && result.success) {
        handleAddMessage({
          role: 'assistant',
          content: result.humanMessage,
          isSuccess: true
        });
      } else if (!result.success) {
        handleAddMessage({
          role: 'assistant',
          content: result.humanMessage,
          isError: true
        });
      }
      
      return result;
    } catch (error) {
      console.error("Error executing VBA code:", error);
      const errorMessage = error.humanMessage || `Error executing VBA code: ${error.message}`;
      
      // Only show error message if not suppressed
      if (!suppressSuccess) {
        handleAddMessage({
          role: 'assistant',
          content: errorMessage,
          isError: true
        });
      }
      
      return { success: false, error: error.message, humanMessage: errorMessage };
    } finally {
      setIsProcessingAction(false);
    }
  };

  const runOfficeJsCode = async (jsCode, suppressSuccess = false) => {
    setIsProcessingAction(true);
    try {
      console.log("Attempting to execute Office.js code:", jsCode);
      const result = await excelService.executeOfficeJsCode(jsCode);
      
      // Check operations history after execution
      try {
        const operationsHistory = excelService.getOperationsHistory();
        console.log("Operations history after executeOfficeJsCode:", operationsHistory);
        setIsUndoAvailable(operationsHistory && operationsHistory.length > 0);
      } catch (error) {
        console.error("Error checking operations history:", error);
      }
      
      if (!result.success) {
        console.error("Error executing Office.js code:", result.error);
        // Display error message to user
        handleAddMessage({
          role: 'assistant',
          content: result.humanMessage || `Error executing Office.js code: ${result.error}. Please make sure Excel is properly loaded and the code is valid.`,
          isError: true
        });
      } else {
        console.log("Office.js code executed successfully");
        // Only show success message if not suppressed (when called manually)
        if (!suppressSuccess) {
          handleAddMessage({
            role: 'assistant',
            content: result.humanMessage || "Code executed successfully!",
            isSuccess: true
          });
        }
      }
      return result;
    } catch (error) {
      console.error("Exception when executing Office.js code:", error);
      // Display error message to user
      const errorMessage = error.humanMessage || `Error executing Office.js code: ${error.message}. Please make sure Excel is properly loaded and the code is valid.`;
      handleAddMessage({
        role: 'assistant',
        content: errorMessage,
        isError: true
      });
      return { success: false, error: error.message, humanMessage: errorMessage };
    } finally {
      setIsProcessingAction(false);
    }
  };

  const toggleAiMode = (mode) => {
    setAiMode(mode);
  };
  
  // Helper function to render debugging information
  const renderDebugInfo = (message) => {
    // Only render debug info if message contains formula-related content
    const containsFormula = message.content.match(/=([A-Za-z]+\([^)]*\))/i);
    if (!containsFormula) return null;
    
    // Extract potential target cells from the message
    const detectedCells = [];
    
    // Check various patterns
    const inCellMatch = message.content.match(/in\s+cell\s+([A-Z]+[0-9]+)/i);
    const intoCellMatch = message.content.match(/into\s+cell\s+([A-Z]+[0-9]+)/i);
    const cellMatch = message.content.match(/cell\s+([A-Z]+[0-9]+)/i);
    
    if (inCellMatch && inCellMatch[1]) detectedCells.push(inCellMatch[1]);
    if (intoCellMatch && intoCellMatch[1]) detectedCells.push(intoCellMatch[1]);
    if (cellMatch && cellMatch[1]) detectedCells.push(cellMatch[1]);
    
    // Check for explicit cell reference
    const extractedRef = message.extractedCellRef;
    if (extractedRef) detectedCells.push(`From user: ${extractedRef}`);
    
    // If we found cells, display the debug info
    if (detectedCells.length > 0) {
      return (
        <div style={{ 
          marginTop: "8px", 
          fontSize: "12px", 
          color: tokens.colorNeutralForeground3,
          padding: "4px 8px",
          backgroundColor: tokens.colorNeutralBackground3,
          borderRadius: "4px"
        }}>
          <Text>System detected target cells: {detectedCells.join(", ")}</Text>
        </div>
      );
    }
    
    return null;
  };

  // Function to update conversation history
  const updateConversationHistory = (updates) => {
    setConversationHistory(prev => ({
      ...prev,
      ...updates,
      // For array updates, ensure we don't exceed a reasonable size and avoid duplicates
      ...(updates.recentCells ? {
        recentCells: [...new Set([...updates.recentCells, ...prev.recentCells])].slice(0, 10)
      } : {}),
      ...(updates.recentActions ? {
        recentActions: [...updates.recentActions, ...prev.recentActions].slice(0, 10)
      } : {})
    }));
  };

  // Function to get historical context string for AI prompt
  const getHistoricalContext = () => {
    const history = conversationHistory;
    let contextString = "\n\nHistorical Context:";
    
    if (history.recentCells.length > 0) {
      contextString += `\n- Recently referenced cells/ranges: ${history.recentCells.join(", ")}`;
    }
    
    if (history.recentActions.length > 0) {
      contextString += "\n- Recent actions performed:";
      history.recentActions.slice(0, 5).forEach((action, index) => {
        contextString += `\n  ${index + 1}. ${action}`;
      });
    }
    
    if (history.lastActiveSheet) {
      contextString += `\n- Last active worksheet: ${history.lastActiveSheet}`;
    }
    
    return contextString;
  };

  // New function to render context history for the user
  const renderContextHistory = () => {
    const history = conversationHistory;
    
    if (!history.recentCells.length && !history.recentActions.length) {
      return null;
    }
    
    return (
      <div style={{
        fontSize: tokens.fontSizeBase200,
        color: tokens.colorNeutralForeground3,
        padding: "8px",
        margin: "8px 0",
        backgroundColor: tokens.colorNeutralBackground2,
        borderRadius: tokens.borderRadiusMedium,
        borderLeft: `3px solid ${tokens.colorBrandStroke1}`
      }}>
        <Text weight="semibold">Context Awareness</Text>
        
        {history.recentCells.length > 0 && (
          <div style={{ marginTop: "4px" }}>
            <Text size={100}>Recent cells: {history.recentCells.slice(0, 5).join(", ")}</Text>
          </div>
        )}
        
        {history.lastActiveSheet && (
          <div style={{ marginTop: "4px" }}>
            <Text size={100}>Active sheet: {history.lastActiveSheet}</Text>
          </div>
        )}
        
        {history.recentActions.length > 0 && (
          <div style={{ marginTop: "4px" }}>
            <Text size={100}>Last action: {history.recentActions[0]}</Text>
          </div>
        )}
      </div>
    );
  };

  // Function to handle getting all data from the active worksheet
  const handleGetAllData = async () => {
    setIsLoading(true);
    try {
      const result = await excelService.getAllData();
      
      if (result.success) {
        const { values, rowCount, columnCount, address } = result;
        
        // Format the data for display
        let formattedData = '';
        
        if (values && values.length > 0) {
          // Create a markdown table with the data
          formattedData = '| ' + Array(columnCount).fill('').map((_, i) => `Column ${i+1}`).join(' | ') + ' |\n';
          formattedData += '| ' + Array(columnCount).fill('---').join(' | ') + ' |\n';
          
          values.forEach(row => {
            formattedData += '| ' + row.map(cell => cell !== undefined ? String(cell) : '').join(' | ') + ' |\n';
          });
        } else {
          formattedData = "The worksheet appears to be empty.";
        }
        
        const summary = `
Data retrieved from the active worksheet:
- Range: ${address}
- Rows: ${rowCount}
- Columns: ${columnCount}

${formattedData}
`;
        
        // Add the retrieved data to the chat
        handleAddMessage({
          role: 'assistant',
          content: summary
        });
      } else {
        handleAddMessage({
          role: 'assistant',
          content: `Error retrieving worksheet data: ${result.error}`
        });
      }
    } catch (error) {
      console.error('Error in handleGetAllData:', error);
      handleAddMessage({
        role: 'assistant',
        content: `Failed to retrieve worksheet data: ${error.message}`
      });
    } finally {
      setIsLoading(false);
    }
  };

  // Helper function to add a message to the chat
  const handleAddMessage = (message) => {
    // If the message is from the assistant and is ONLY a code block, do not display the code, just execute and show feedback
    if (
      message.role === 'assistant' &&
      typeof message.content === 'string' &&
      message.content.trim().match(/^```(js|javascript)\s*[\s\S]*?```\s*$/) &&
      aiMode !== 'PROMPT'
    ) {
      // Extract the code
      const codeMatch = message.content.trim().match(/^```(js|javascript)\s*([\s\S]*?)```\s*$/);
      if (codeMatch && codeMatch[2]) {
        const jsCode = codeMatch[2].trim();
        
        // Extract a human-friendly description of what the code does
        const operationType = extractOperationType(jsCode);
        const friendlyMessage = generateFriendlyMessage(operationType, jsCode);
        
        // Add a friendly message to the chat
            setMessages(prev => [...prev, {
              id: Date.now(),
              role: 'assistant',
          content: friendlyMessage,
              isSuccess: true
            }]);
        
        // Then execute the code
        runOfficeJsCode(jsCode, true).then((result) => {
          if (!result || !result.success) {
            // Only add an error message if execution failed
            setMessages(prev => [...prev, {
              id: Date.now(),
              role: 'assistant',
              content: `Sorry, there was an error: ${result && result.error ? result.error : 'Unknown error.'}`,
              isError: true
            }]);
          }
        });
        return;
      }
    }
    // Otherwise, show the message as usual
    const newMessage = {
      id: Date.now(),
      ...message
    };
    setMessages(prev => [...prev, newMessage]);
  };

  return (
    <div className={styles.container}>
      <div className={styles.chatHeader}>
        <div className={styles.headerLeft}>
          <Bot24Regular />
          <Text>Excel AI Assistant</Text>
        </div>
        
        <div className={styles.headerRight}>
          <Tooltip content="Undo last operation" relationship="label">
            <Button 
              className={styles.undoButton}
              icon={<ArrowUndo24Regular />}
              onClick={handleUndo}
              disabled={!isUndoAvailable || isUndoInProgress}
              appearance="subtle"
              size="small"
            >
              Undo {isUndoAvailable ? 'Available' : 'Unavailable'}
            </Button>
          </Tooltip>
        
          {/* Hidden test button for development */}
          <Button
            size="small"
            appearance="subtle"
            onClick={handleTestUndo}
            style={{ fontSize: '10px', padding: '2px 4px' }}
          >
            Test Undo
          </Button>
        
        <Menu>
          <MenuTrigger disableButtonEnhancement>
            <div className={styles.modeSelector}>
              <AppsAddIn24Regular />
              <Text>{AI_MODES[aiMode].name}</Text>
              <ChevronDown20Regular />
            </div>
          </MenuTrigger>
          <MenuPopover>
            <MenuList>
              <MenuItem 
                onClick={() => toggleAiMode("ASK")}
                icon={aiMode === "ASK" ? <Badge appearance="filled" /> : null}
              >
                {AI_MODES.ASK.name}
                <Text size={100} block>{AI_MODES.ASK.description}</Text>
              </MenuItem>
              <MenuItem 
                onClick={() => toggleAiMode("AGENT")}
                icon={aiMode === "AGENT" ? <Badge appearance="filled" /> : null}
              >
                {AI_MODES.AGENT.name}
                <Text size={100} block>{AI_MODES.AGENT.description}</Text>
              </MenuItem>
              <MenuItem 
                onClick={() => toggleAiMode("PROMPT")}
                icon={aiMode === "PROMPT" ? <Badge appearance="filled" /> : null}
              >
                {AI_MODES.PROMPT.name}
                <Text size={100} block>{AI_MODES.PROMPT.description}</Text>
              </MenuItem>
            </MenuList>
          </MenuPopover>
        </Menu>
        </div>
      </div>

      {error && (
        <MessageBar intent="error" className={styles.errorMessage}>
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}

      {!isApiKeyValid && (
        <MessageBar intent="warning">
          <MessageBarBody>Please set your OpenAI API key to use the AI features.</MessageBarBody>
        </MessageBar>
      )}

      {/* Add context history display */}
      {messages.length > 0 && renderContextHistory()}

      <div className={styles.chatContainer} ref={chatContainerRef}>
        {messages.length === 0 ? (
          <div className={styles.emptyChatMessage}>
            <Bot24Regular className={styles.robotIcon} />
            <Text size={500} weight="semibold">How can I help with your Excel tasks?</Text>
            <Text size={300}>Ask me about formulas, data analysis, or Excel functions</Text>
            
            <div className={styles.suggestions}>
              {EXCEL_SUGGESTIONS.map((suggestion, index) => (
                <Button
                  key={index}
                  appearance="subtle"
                  size="medium"
                  icon={suggestion.icon}
                  className={styles.suggestionButton}
                  onClick={() => handleSuggestionClick(suggestion.text)}
                >
                  {suggestion.text}
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
        ) : (
          messages.map((message) => (
            <div
              key={message.id}
              className={`${styles.message} ${
                message.role === "user" ? styles.userMessage : styles.aiMessage
              }`}
            >
              {message.role === "user" && (
                <Tooltip content="Resend this message" relationship="label">
                  <div 
                    className={styles.resendButton}
                    onClick={() => resendMessage(message.content)}
                    aria-label="Resend message"
                  >
                    <ArrowRotateClockwise20Regular />
                  </div>
                </Tooltip>
              )}
              <Card 
                className={message.role === "user" ? styles.userCard : styles.aiCard}
              >
                <CardHeader
                  image={
                    <Avatar 
                      className={styles.avatar}
                      icon={message.role === "user" ? <Person24Regular /> : <Bot24Regular />} 
                      color={message.role === "user" ? "neutral" : "brand"}
                    />
                  }
                  header={
                    <Text weight="semibold">
                      {message.role === "user" ? "You" : "Excel AI Assistant"}
                    </Text>
                  }
                />
                <CardPreview>
                  {renderMessageContent(message)}
                </CardPreview>
              </Card>
            </div>
          ))
        )}
        
        {isLoading && (
          <div className={styles.loadingContainer}>
            <Spinner size="tiny" />
            <Text>Thinking...</Text>
          </div>
        )}
      </div>

      <div className={styles.messageInput}>
        <Input
          placeholder="Ask about Excel formulas, data analysis, or help with your spreadsheet..."
          value={inputValue}
          onChange={(e) => setInputValue(e.target.value)}
          onKeyDown={handleKeyDown}
          contentAfter={
            <Button
              appearance={isResendMode ? "primary" : "transparent"}
              icon={<Send24Regular />}
              onClick={handleSendMessage}
              disabled={!inputValue.trim() || isLoading || !isApiKeyValid}
            />
          }
          disabled={!isApiKeyValid}
          style={{ 
            width: '100%',
            border: isResendMode ? `1px solid ${tokens.colorBrandBorder1}` : undefined,
            backgroundColor: isResendMode ? tokens.colorBrandBackground1 : undefined
          }}
          size="large"
        />
      </div>
    </div>
  );
};

export default AIChat; 