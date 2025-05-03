import React, { useState, useEffect, useRef } from 'react';
import {
  makeStyles,
  tokens,
  Spinner,
  Text,
  MessageBar,
  MessageBarBody,
  Title3,
  Body1,
  Card,
  CardHeader,
  Button,
  mergeClasses,
  Badge,
  Tab,
  TabList,
} from "@fluentui/react-components";
import { useAI } from '../../context/AIContext';
import ChatMessage from './chat/ChatMessage';
import ChatInput from './chat/ChatInput';
import ChatHeader from './chat/ChatHeader';
import SuggestionsList from './chat/SuggestionsList';
import EmptyChat from './chat/EmptyChat';
import { v4 as uuidv4 } from 'uuid';
import {
  Dismiss24Regular, 
  DataArea24Regular,
  TableSimple24Regular,
  InfoRegular,
  ChevronDown20Regular,
  ChevronUp20Regular,
  Copy24Regular,
  ArrowDownload24Regular,
  MoreHorizontal20Regular,
  Table24Regular
} from "@fluentui/react-icons";

// Define AI operation modes with descriptions
export const AI_MODES = {
  ASK: { key: 'ASK', text: 'Ask AI', description: 'Ask questions about your Excel data or get general Excel help' },
  AGENT: { key: 'AGENT', text: 'AI Agent', description: 'Let AI analyze your data and take actions on your behalf' },
  PROMPT: { key: 'PROMPT', text: 'Prompt Library', description: 'Use specialized prompts for common Excel tasks' }
};

// Sample suggestions for Excel tasks
const SUGGESTIONS = {
  ASK: [
    'Explain the data in my selected range',
    'How do I calculate a running total?',
    'What formula should I use to count unique values?',
    'Help me understand pivot tables'
  ],
  AGENT: [
    'Clean and format this data',
    'Analyze this data and create a summary',
    'Find patterns or outliers in this data',
    'Create a chart that best shows this data'
  ],
  PROMPT: [
    'Create a sales dashboard',
    'Format this table professionally',
    'Add data validation to these cells',
    'Create a dynamic named range'
  ]
};

// Custom styles for the chat component
const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    height: "100%",
    padding: "16px",
    background: tokens.colorNeutralBackground2,
    boxSizing: "border-box"
  },
  chatContainer: {
    display: "flex",
    flexDirection: "column",
    flexGrow: 1,
    overflow: "hidden",
    boxShadow: `${tokens.shadow16}, 0 0 0 1px ${tokens.colorNeutralStroke1}`,
    borderRadius: tokens.borderRadiusLarge,
    backgroundColor: tokens.colorNeutralBackground1
  },
  messagesContainer: {
    flexGrow: 1,
    overflow: "auto",
    padding: "20px 24px",
    display: "flex",
    flexDirection: "column",
    gap: "16px",
    backgroundImage: `radial-gradient(${tokens.colorNeutralBackground2} 1px, transparent 1px)`,
    backgroundSize: "24px 24px",
    backgroundPosition: "-12px -12px",
    "& > div:nth-child(n)": {
      animationDelay: "calc(0.05s * var(--message-index, 0))",
    }
  },
  apiKeyError: {
    marginBottom: "12px"
  },
  loadingOverlay: {
    position: "absolute",
    top: 0,
    left: 0,
    right: 0,
    bottom: 0,
    backgroundColor: "rgba(255, 255, 255, 0.7)",
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
    zIndex: 2,
    flexDirection: "column"
  },
  loadingText: {
    marginTop: "8px",
    fontSize: "14px",
    color: tokens.colorNeutralForeground1
  },
  dataViewButton: {
    position: "fixed",
    bottom: "90px",
    right: "24px",
    zIndex: 100,
    borderRadius: "50%",
    width: "48px",
    height: "48px",
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForeground1BrandSelected,
    boxShadow: tokens.shadow16,
    transition: "all 0.3s ease",
    border: `2px solid ${tokens.colorBrandBackgroundHover}`,
    cursor: "pointer",
    padding: "0",
    minWidth: "unset",
    "&:hover": {
      transform: "scale(1.08)",
      backgroundColor: tokens.colorBrandBackgroundHover,
      boxShadow: "0 6px 16px rgba(0, 0, 0, 0.2)",
    },
    "&:active": {
      transform: "scale(0.95)",
    },
    "& svg": {
      width: "24px",
      height: "24px",
    }
  },
  
  dataButtonWrapper: {
    position: "relative",
    display: "inline-block"
  },
  dataBadge: {
    position: "absolute",
    top: "-8px",
    right: "-8px",
    backgroundColor: tokens.colorStatusDangerBackground,
    color: tokens.colorStatusDangerForeground1,
    minWidth: "20px",
    height: "20px",
    borderRadius: "10px",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    fontSize: "11px",
    fontWeight: "bold",
    padding: "0 6px",
    boxShadow: tokens.shadow4,
    border: `1px solid ${tokens.colorStatusDangerForeground1}`,
    zIndex: 101
  },
  dataButtonLabel: {
    position: "absolute",
    bottom: "-20px",
    left: "50%",
    transform: "translateX(-50%)",
    whiteSpace: "nowrap",
    fontSize: "10px",
    backgroundColor: tokens.colorNeutralBackground1,
    padding: "2px 6px",
    borderRadius: "4px",
    boxShadow: tokens.shadow4,
    opacity: 0.9
  },
  dataDialog: {
    maxWidth: "90vw",
    maxHeight: "80vh",
    width: "min(600px, 90vw)",
    position: "fixed", 
    top: "50%",
    left: "50%",
    transform: "translate(-50%, -50%)",
    margin: 0,
    borderRadius: "12px",
    boxShadow: "0 8px 32px rgba(0, 0, 0, 0.12)",
    border: `1px solid ${tokens.colorNeutralStroke1}`
  },
  dataDialogContent: {
    overflowY: "auto",
    maxHeight: "calc(70vh - 120px)", // Adjust based on header and footer height
    padding: "4px 12px 8px 12px",
    "&::-webkit-scrollbar": {
      width: "8px",
      height: "8px"
    },
    "&::-webkit-scrollbar-thumb": {
      backgroundColor: tokens.colorNeutralStroke3,
      borderRadius: "4px"
    },
    "&::-webkit-scrollbar-track": {
      backgroundColor: tokens.colorNeutralBackground2,
      borderRadius: "4px"
    }
  },
  codeBlock: {
    backgroundColor: tokens.colorNeutralBackground3,
    padding: "12px",
    borderRadius: "8px",
    fontFamily: "Consolas, Monaco, 'Andale Mono', monospace",
    overflowX: "auto",
    fontSize: tokens.fontSizeBase200,
    lineHeight: tokens.lineHeightBase300,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    margin: "12px 0",
    whiteSpace: "pre-wrap",
    position: "relative"
  },
  codeHeader: {
    position: "absolute",
    top: "8px",
    right: "8px",
    display: "flex",
    gap: "8px"
  },
  dataPreviewTable: {
    width: "100%",
    borderCollapse: "collapse",
    fontSize: tokens.fontSizeBase200,
    marginTop: "15px",
    borderRadius: "4px",
    overflow: "hidden", // For the rounded corners
    border: `1px solid ${tokens.colorNeutralStroke2}`
  },
  dataTableHeader: {
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForeground1BrandSelected,
    fontWeight: tokens.fontWeightSemibold,
    textAlign: "left",
    padding: "10px 12px",
    position: "sticky",
    top: 0,
    zIndex: 1,
    whiteSpace: "nowrap",
    "&:first-child": {
      borderTopLeftRadius: "4px"
    },
    "&:last-child": {
      borderTopRightRadius: "4px"
    }
  },
  dataTableCell: {
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderWidth: "0 0 1px 0",
    padding: "8px 12px",
    maxWidth: "200px",
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap"
  },
  dataTableRow: {
    "&:nth-child(even)": {
      backgroundColor: tokens.colorNeutralBackground2
    },
    "&:hover": {
      backgroundColor: tokens.colorNeutralBackground3
    },
    "&:last-child td": {
      borderBottomWidth: "0"
    }
  },
  dataActionButton: {
    padding: "4px 8px",
    marginRight: "8px"
  },
  expandCollapseButton: {
    backgroundColor: "transparent",
    border: "none",
    cursor: "pointer",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    padding: "4px",
    borderRadius: "4px",
    transition: "all 0.2s ease",
    "&:hover": {
      backgroundColor: tokens.colorNeutralBackground3
    }
  },
  fullSizePreview: {
    maxHeight: "none",
    height: "auto"
  },
  collapseToggleButton: {
    margin: "16px auto",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    gap: "8px",
    padding: "6px 16px",
    borderRadius: "16px",
    backgroundColor: tokens.colorNeutralBackground3,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    cursor: "pointer",
    transition: "all 0.2s ease",
    "&:hover": {
      backgroundColor: tokens.colorNeutralBackground4,
      transform: "translateY(-1px)"
    }
  },
  dialogHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    paddingBottom: "12px",
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    marginBottom: "12px"
  },
  closeButton: {
    minWidth: "32px",
    width: "32px",
    height: "32px",
    borderRadius: "16px",
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
    padding: 0,
    color: tokens.colorNeutralForeground3,
    "&:hover": {
      backgroundColor: tokens.colorNeutralBackground3,
      color: tokens.colorNeutralForeground1
    }
  },
  dataInfo: {
    margin: "12px 0",
    padding: "12px",
    backgroundColor: tokens.colorBrandBackground2,
    borderRadius: "8px",
    border: `1px solid ${tokens.colorBrandStroke2}`,
  },
  dialogTitle: {
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1
  },
  tabList: {
    marginBottom: "12px"
  },
  tableWrapper: {
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: "4px",
    overflow: "hidden",
    marginTop: "16px",
    marginBottom: "16px",
    boxShadow: "0 2px 4px rgba(0, 0, 0, 0.05)"
  },
  fullScreenOverlay: {
    position: "absolute",
    top: 0,
    left: 0,
    right: 0,
    bottom: 0,
    backgroundColor: tokens.colorNeutralBackground1,
    zIndex: 1000,
    display: "flex",
    flexDirection: "column",
    overflow: "hidden",
    borderRadius: tokens.borderRadiusLarge,
    boxShadow: tokens.shadow64
  },
  
  dataViewerHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    padding: "16px 24px",
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1
  },
  
  dataViewerTitle: {
    fontSize: tokens.fontSizeBase600,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1
  },
  
  dataViewerContent: {
    flexGrow: 1,
    overflowY: "auto",
    padding: "16px 24px",
    "&::-webkit-scrollbar": {
      width: "8px",
      height: "8px"
    },
    "&::-webkit-scrollbar-thumb": {
      backgroundColor: tokens.colorNeutralStroke3,
      borderRadius: "4px"
    },
    "&::-webkit-scrollbar-track": {
      backgroundColor: tokens.colorNeutralBackground2,
      borderRadius: "4px"
    }
  },
  
  dataViewerFooter: {
    display: "flex",
    justifyContent: "flex-end",
    padding: "16px 24px",
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
    gap: "8px"
  },
  
  statusMessage: {
    position: "fixed",
    bottom: "140px",
    left: "50%",
    transform: "translateX(-50%)",
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForeground1BrandSelected,
    borderRadius: "16px",
    padding: "8px 16px",
    boxShadow: tokens.shadow16,
    zIndex: 1000,
    animation: "fadeInOut 0.3s ease",
    maxWidth: "80%",
    textAlign: "center",
    fontSize: tokens.fontSizeBase200
  },
  
  "@keyframes fadeInOut": {
    "0%": { opacity: 0, transform: "translate(-50%, 20px)" },
    "100%": { opacity: 1, transform: "translate(-50%, 0)" }
  },
});

// View Data Button Component
const ViewDataButton = ({ selectionData, formatDataForDisplay, hasSelection, openAIJsonData }) => {
  const styles = useStyles();
  const [open, setOpen] = React.useState(false);
  const [showLabel, setShowLabel] = React.useState(false);
  const [expanded, setExpanded] = React.useState(false);
  const [viewMode, setViewMode] = React.useState('table'); // 'table', 'raw', 'json', or 'openai'
  
  // Only render if there's a selection
  if (!hasSelection) return null;
  
  // Format the data for display
  const { values, hasHeaders, address, rowCount, columnCount } = selectionData || {};
  
  // Raw data formatting (same as before)
  const formattedData = values ? formatDataForDisplay(values, hasHeaders) : "No data available";
  
  // For row count in badge
  const rowCount1 = rowCount || (values ? values.length : 0);
  const badgeCount = rowCount1 > 0 ? rowCount1 : null;
  
  // Format large numbers with K suffix
  const formatBadgeCount = (count) => {
    if (!count) return "";
    if (count > 999) {
      return (count / 1000).toFixed(1) + "K";
    }
    return count;
  };
  
  // Format as JSON for alternate view
  const getJsonData = () => {
    if (!values || !Array.isArray(values) || values.length === 0) {
      return "[]";
    }
    
    try {
      // If first row contains headers, use it to create objects
      if (hasHeaders && values.length > 1) {
        const headers = values[0];
        const jsonData = values.slice(1).map(row => {
          const obj = {};
          headers.forEach((header, index) => {
            obj[header] = row[index];
          });
          return obj;
        });
        return JSON.stringify(jsonData, null, 2);
      } else {
        // If no headers, just return array of arrays
        return JSON.stringify(values, null, 2);
      }
    } catch (err) {
      console.error("Error formatting JSON:", err);
      return "Error creating JSON";
    }
  };
  
  // Get the OpenAI JSON data
  const getOpenAIJsonData = () => {
    console.log("getOpenAIJsonData called, openAIJsonData:", openAIJsonData ? "exists" : "null");
    
    if (!openAIJsonData) {
      // Check if we can create data from selectionData
      if (selectionData && selectionData.values) {
        try {
          console.log("Creating display data from selectionData");
          const values = selectionData.values;
          const hasHeadersLocal = selectionData.hasHeaders || false;
          
          // Create headers
          const headers = hasHeadersLocal && values.length > 0 ? 
            values[0].map((h, idx) => h || `Column_${idx+1}`) : 
            Array.from({length: values[0]?.length || 0}, (_, i) => `Column_${i+1}`);
          
          // Create simple JSON data
          const startRow = hasHeadersLocal ? 1 : 0;
          const jsonData = [];
          
          for (let r = startRow; r < values.length; r++) {
            const row = {};
            for (let c = 0; c < values[r].length; c++) {
              row[headers[c]] = values[r][c];
            }
            jsonData.push(row);
          }
          
          // Create a display version
          const displayData = {
            note: "⚠️ DATA PREVIEW ONLY: This data has NOT been sent to OpenAI yet. You must ask a question first.",
            headers,
            dataRows: jsonData,
            rowCount: jsonData.length,
            columnCount: headers.length,
          };
          
          return JSON.stringify(displayData, null, 2);
        } catch (err) {
          console.error("Error creating display data:", err);
        }
      }
      
      // If we can't create data, show the default message
      console.log("No openAIJsonData available in viewer");
      return "No data has been sent to OpenAI yet. Try asking a question about your data first.";
    }
    
    try {
      console.log("Formatting OpenAI data for display", 
        openAIJsonData.jsonData?.length || 0, "rows");
      
      // Format the main data and important metadata
      const formattedOpenAIData = {
        note: openAIJsonData.source === "enrichedContext" 
          ? "✅ EXACT DATA: This is the exact JSON data that was sent to OpenAI for analysis"
          : openAIJsonData.source === "directInjection"
            ? "✅ DIRECT INJECTION: This data was directly injected into the prompt for OpenAI analysis"
            : "⚠️ RECONSTRUCTED DATA: This data was reconstructed from your selection and may differ from what OpenAI received",
        timestamp: openAIJsonData.timestamp || new Date().toISOString(),
        dataRows: openAIJsonData.jsonData,
        rowCount: openAIJsonData.jsonData?.length || 0,
        headers: openAIJsonData.headers || [],
        columnCount: openAIJsonData.headers?.length || 0
      };
      
      // Add a sample of the prompt (first 200 chars) to show what instructions were given
      if (openAIJsonData.fullPrompt) {
        formattedOpenAIData.promptSample = openAIJsonData.fullPrompt.substring(0, 200) + "...";
      }
      
      return JSON.stringify(formattedOpenAIData, null, 2);
    } catch (err) {
      console.error("Error formatting OpenAI JSON:", err);
      return "Error rendering OpenAI data format: " + err.message;
    }
  };
  
  // Create a formatted HTML table 
  const renderTablePreview = () => {
    if (!values || !Array.isArray(values) || values.length === 0) {
      return <Text>No data available</Text>;
    }
    
    // Determine row limit based on expanded state
    const rowLimit = expanded ? values.length : Math.min(10, values.length);
    const showExpander = values.length > 10;
    
    return (
      <div>
        <div className={styles.tableWrapper}>
          <div style={{ overflowX: "auto" }}>
            <table className={styles.dataPreviewTable}>
              {hasHeaders && values.length > 1 && (
                <thead>
                  <tr>
                    {values[0].map((header, index) => (
                      <th key={index} className={styles.dataTableHeader}>
                        {header !== null && header !== undefined ? String(header) : `Column ${index+1}`}
                      </th>
                    ))}
                  </tr>
                </thead>
              )}
              <tbody>
                {values.slice(hasHeaders ? 1 : 0, rowLimit).map((row, rowIndex) => (
                  <tr key={rowIndex} className={styles.dataTableRow}>
                    {row.map((cell, cellIndex) => (
                      <td key={cellIndex} className={styles.dataTableCell} title={cell !== null ? String(cell) : ''}>
                        {cell !== null && cell !== undefined ? String(cell) : ''}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
        
        {showExpander && (
          <button 
            className={styles.collapseToggleButton} 
            onClick={() => setExpanded(!expanded)}
            aria-label={expanded ? "Show less rows" : "Show all rows"}
          >
            {expanded ? (
              <>
                <ChevronUp20Regular />
                <Text>Show fewer rows</Text>
              </>
            ) : (
              <>
                <ChevronDown20Regular />
                <Text>Show all {values.length - (hasHeaders ? 1 : 0)} rows</Text>
              </>
            )}
          </button>
        )}
      </div>
    );
  };
  
  // Download data as CSV
  const downloadCsv = () => {
    if (!values) return;
    
    try {
      const csv = values.map(row => 
        row.map(cell => {
          // Handle special characters in CSV
          if (cell === null || cell === undefined) return '';
          const cellStr = String(cell);
          // Escape quotes and wrap in quotes if contains commas, quotes or newlines
          if (cellStr.includes(',') || cellStr.includes('"') || cellStr.includes('\n')) {
            return `"${cellStr.replace(/"/g, '""')}"`;
          }
          return cellStr;
        }).join(",")
      ).join("\n");
      
      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.setAttribute('href', url);
      link.setAttribute('download', `excel-data-${new Date().toISOString().slice(0,10)}.csv`);
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    } catch (err) {
      console.error("Error downloading CSV:", err);
    }
  };
  
  return (
    <>
      <div 
        className={styles.dataButtonWrapper} 
        onMouseEnter={() => setShowLabel(true)} 
        onMouseLeave={() => setShowLabel(false)}
        onClick={() => setOpen(true)}
      >
        <Button 
          className={styles.dataViewButton}
          title="View Selected Data"
          aria-label="View Selected Data"
        >
          <Table24Regular />
        </Button>
        {badgeCount && (
          <span className={styles.dataBadge}>
            {formatBadgeCount(badgeCount)}
          </span>
        )}
        {showLabel && (
          <span className={styles.dataButtonLabel}>
            View Excel Data
          </span>
        )}
      </div>

      {open && (
        <div className={styles.fullScreenOverlay}>
          <div className={styles.dataViewerHeader}>
            <Text className={styles.dataViewerTitle}>Excel Data Selection</Text>
            <Button 
              appearance="subtle" 
              icon={<Dismiss24Regular />}
              className={styles.closeButton}
              onClick={() => setOpen(false)}
              aria-label="Close data viewer"
            />
          </div>
          
          <div className={styles.dataViewerContent}>
            <div className={styles.dataInfo}>
              <Text weight="semibold">📊 Selection Information</Text>
              <Text block>Range: {address || 'Unknown'}</Text>
              <Text block>Dimensions: {rowCount || 0} rows × {columnCount || 0} columns</Text>
              {hasHeaders && <Text block>First row contains headers</Text>}
            </div>
            
            <TabList 
              selectedValue={viewMode} 
              onTabSelect={(e, data) => setViewMode(data.value)}
              className={styles.tabList}
            >
              <Tab value="table">Table View</Tab>
              <Tab value="raw">Raw Data</Tab>
              <Tab value="json">Simple JSON</Tab>
              <Tab value="openai">OpenAI Data</Tab>
            </TabList>
            
            {viewMode === 'table' && renderTablePreview()}
            
            {viewMode === 'raw' && (
              <div className={styles.codeBlock}>
                <div className={styles.codeHeader}>
                  <Button 
                    icon={<Copy24Regular />}
                    appearance="subtle"
                    size="small"
                    onClick={() => {
                      if (values) {
                        const text = values.map(row => row.join("\t")).join("\n");
                        navigator.clipboard.writeText(text);
                      }
                    }}
                    aria-label="Copy raw data"
                  >
                    Copy
                  </Button>
                </div>
                {formattedData}
              </div>
            )}
            
            {viewMode === 'json' && (
              <div className={styles.codeBlock}>
                <div className={styles.codeHeader}>
                  <Button 
                    icon={<Copy24Regular />}
                    appearance="subtle"
                    size="small"
                    onClick={() => navigator.clipboard.writeText(getJsonData())}
                    aria-label="Copy JSON"
                  >
                    Copy
                  </Button>
                </div>
                <pre>{getJsonData()}</pre>
              </div>
            )}
            
            {viewMode === 'openai' && (
              <div className={styles.codeBlock}>
                <div className={styles.codeHeader}>
                  <Button 
                    icon={<Copy24Regular />}
                    appearance="subtle"
                    size="small"
                    onClick={() => navigator.clipboard.writeText(getOpenAIJsonData())}
                    aria-label="Copy OpenAI JSON"
                  >
                    Copy
                  </Button>
                </div>
                <div>
                  <Text weight="semibold" block>This is the exact JSON data sent to OpenAI for analysis:</Text>
                  <Text block size="small" style={{ marginBottom: '8px', color: tokens.colorNeutralForeground2 }}>
                    The data includes preserved types, column metadata, and verification information.
                  </Text>
                </div>
                <pre>{getOpenAIJsonData()}</pre>
              </div>
            )}
          </div>
          
          <div className={styles.dataViewerFooter}>
            <Button appearance="secondary" onClick={() => setOpen(false)}>Close</Button>
            <Button 
              appearance="primary" 
              icon={<ArrowDownload24Regular />}
              onClick={downloadCsv}
            >
              Download CSV
            </Button>
          </div>
        </div>
      )}
    </>
  );
};

// Main chat component
export default function AIChat() {
  const classes = useStyles();
  const { 
    apiKey, 
    isProcessing,
    error,
    messages,
    addMessage,
    updateLastMessage,
    clearMessages,
    generateText,
    executeExcelCode,
    canUndo,
    undoLastOperation,
    loadExcelContext,
    analyzeData,
    suggestFormula,
    suggestChart,
    loadDataProfile,
    getSelectionData,
    countValueInColumn,
    countValueInRange,
    analyzeColumnData,
    answerDataQuestion,
    excelService,
    getExcelContext,
    abortAnalysis,
    checkApiKey
  } = useAI();
  
  // Local state
  const [inputValue, setInputValue] = useState('');
  const [processingAction, setProcessingAction] = useState(false);
  const [mode, setMode] = useState(AI_MODES.ASK);
  const [actionText, setActionText] = useState('');
  const [excelDataLoaded, setExcelDataLoaded] = useState(false);
  const [suggestionsVisible, setSuggestionsVisible] = useState(true);
  const [currentSelection, setCurrentSelection] = useState(null);
  const [useSelection, setUseSelection] = useState(false);
  const [selectionData, setSelectionData] = useState(null);
  const [openAIJsonData, setOpenAIJsonData] = useState(null); // Store OpenAI JSON data
  const [loading, setLoading] = useState(false);
  const [statusMessage, setStatusMessage] = useState(null);
  
  // References
  const messagesEndRef = useRef(null);
  const chatContainerRef = useRef(null);
  
  // Scroll to bottom when messages change
  useEffect(() => {
    if (messagesEndRef.current) {
      messagesEndRef.current.scrollIntoView({ behavior: 'smooth' });
    }
  }, [messages]);

  // Load Excel data before sending messages
  useEffect(() => {
    async function loadData() {
      try {
        const context = await loadExcelContext();
        setExcelDataLoaded(!!context);
        
        // Save selection if it exists
        if (context && context.selection && context.selection.address) {
          setCurrentSelection(context.selection);
        }
      } catch (err) {
        console.error('Error loading Excel data:', err);
        setExcelDataLoaded(false);
      }
    }
    
    loadData();
    
    // Set up an interval to refresh the selection
    const selectionInterval = setInterval(async () => {
      if (useSelection) {
        try {
          const context = await loadExcelContext(true);
          if (context && context.selection && context.selection.address) {
            // Only update if the selection address has changed
            if (!currentSelection || 
                context.selection.address !== currentSelection.address ||
                context.selection.rowCount !== currentSelection.rowCount ||
                context.selection.columnCount !== currentSelection.columnCount) {
              console.log('Selection changed, updating selection info');
            setCurrentSelection(context.selection);
            }
          }
        } catch (err) {
          console.error('Error refreshing selection:', err);
        }
      }
    }, 2000); // Check every 2 seconds
    
    return () => clearInterval(selectionInterval);
  }, [loadExcelContext, useSelection, currentSelection]);
  
  /**
   * Toggle using the current selection for analysis
   */
  const handleToggleSelection = async () => {
    try {
      if (!useSelection) {
        // Turn on selection mode and get current selection
        const context = await loadExcelContext(true);
        if (context && context.selection && context.selection.address) {
          setCurrentSelection(context.selection);
          setUseSelection(true);
        } else {
          // No active selection, show message
          addMessage({
            role: 'assistant',
            content: 'Please select a range in Excel before using this feature.',
            isInfo: true,
            timestamp: new Date().toISOString()
          });
        }
      } else {
        // Turn off selection mode
        setUseSelection(false);
      }
    } catch (err) {
      console.error('Error toggling selection mode:', err);
    }
  };
  
  /**
   * Format data for display in the chat
   * @param {Array} values - 2D array of data
   * @param {boolean} hasHeaders - Whether the data has headers
   * @returns {string} Formatted data for display
   */
  const formatDataForDisplay = (values, hasHeaders) => {
    if (!values || !Array.isArray(values) || values.length === 0) {
      return "No data available";
    }

    // Determine if this is a small or large dataset
    const rowCount = values.length;
    const columnCount = values[0]?.length || 0;
    const isLargeDataset = rowCount > 10 || columnCount > 6;
    
    // For large datasets, show a preview instead of all data
    const showPreview = isLargeDataset;
    
    // Decide how many rows to show
    let rowsToShow = showPreview ? Math.min(6, rowCount) : rowCount;
    let colsToShow = showPreview ? Math.min(6, columnCount) : columnCount;
    
    // For previews, ensure we show at least the first 3 rows if available
    // and some additional rows from the end if there are more than 6 rows
    let previewRows = [];
    if (showPreview && rowCount > 6) {
      // Add first rows up to 3
      previewRows = values.slice(0, 3);
      
      // Add last rows to total 6 rows
      if (rowCount > 6) {
        previewRows.push([...Array(columnCount)].map(() => '...'));
        previewRows = previewRows.concat(values.slice(Math.max(rowCount - 3, 3), rowCount));
      }
    } else {
      previewRows = values.slice(0, rowsToShow);
    }
    
    // Determine which columns to show
    let previewCols = [];
    if (showPreview && columnCount > 6) {
      // Show first 3 and last 3 columns if we have more than 6
      for (let row of previewRows) {
        let newRow = row.slice(0, 3);
        newRow.push('...');
        newRow = newRow.concat(row.slice(Math.max(0, columnCount - 2), columnCount));
        previewCols.push(newRow);
      }
    } else {
      previewCols = previewRows.map(row => row.slice(0, colsToShow));
    }
    
    // Use the preview data for formatting
    const dataToFormat = previewCols;

    let formattedData = "```\n"; // Use code block for formatting
    
    // Format as a table
    const columnWidths = [];
    // Calculate column widths
    for (let c = 0; c < dataToFormat[0].length; c++) {
      let maxWidth = 0;
      for (let r = 0; r < dataToFormat.length; r++) {
        if (!dataToFormat[r]) continue;
        const cellValue = dataToFormat[r][c] !== null && dataToFormat[r][c] !== undefined 
          ? String(dataToFormat[r][c]) 
          : '';
        maxWidth = Math.max(maxWidth, cellValue.length);
      }
      // Cap column width at 20 characters
      columnWidths.push(Math.min(maxWidth, 20));
    }
    
    // Build header row
    if (hasHeaders && dataToFormat.length > 0) {
      let headerRow = "";
      let separatorRow = "";
      
      for (let c = 0; c < dataToFormat[0].length; c++) {
        const cellValue = dataToFormat[0][c] !== null && dataToFormat[0][c] !== undefined 
          ? String(dataToFormat[0][c]) 
          : `Column ${c+1}`;
        const cappedValue = cellValue.length > columnWidths[c] 
          ? cellValue.substring(0, columnWidths[c] - 3) + "..." 
          : cellValue;
        headerRow += cappedValue.padEnd(columnWidths[c]) + " | ";
        separatorRow += "-".repeat(columnWidths[c]) + "-+-";
      }
      
      formattedData += headerRow.trim() + "\n";
      formattedData += separatorRow.trim() + "\n";
      
      // Build data rows, starting from row 1 (after header)
      for (let r = 1; r < dataToFormat.length; r++) {
        let dataRow = "";
        for (let c = 0; c < dataToFormat[r].length; c++) {
          const cellValue = dataToFormat[r][c] !== null && dataToFormat[r][c] !== undefined 
            ? String(dataToFormat[r][c]) 
            : '';
          const cappedValue = cellValue.length > columnWidths[c] 
            ? cellValue.substring(0, columnWidths[c] - 3) + "..." 
            : cellValue;
          dataRow += cappedValue.padEnd(columnWidths[c]) + " | ";
        }
        formattedData += dataRow.trim() + "\n";
      }
    } else {
      // No headers, just format all rows
      for (let r = 0; r < dataToFormat.length; r++) {
        let dataRow = "";
        for (let c = 0; c < dataToFormat[r].length; c++) {
          const cellValue = dataToFormat[r][c] !== null && dataToFormat[r][c] !== undefined 
            ? String(dataToFormat[r][c]) 
            : '';
          const cappedValue = cellValue.length > columnWidths[c] 
            ? cellValue.substring(0, columnWidths[c] - 3) + "..." 
            : cellValue;
          dataRow += cappedValue.padEnd(columnWidths[c]) + " | ";
        }
        formattedData += dataRow.trim() + "\n";
      }
    }
    
    formattedData += "```\n";
    
    // Add summary and data information
    formattedData += `\n📊 **Dataset Information:**\n`;
    formattedData += `- Range: **${currentSelection?.address || 'unknown'}**\n`;
    formattedData += `- Dimensions: **${rowCount} rows × ${columnCount} columns**\n`;
    
    if (showPreview) {
      formattedData += `- Showing preview of ${isLargeDataset ? 'large dataset' : 'data'}\n`;
    }
    
    // Add data type detection
    if (values.length > 1) {
      // Try to detect some data types in first data row (not header)
      const dataRowIndex = hasHeaders ? 1 : 0;
      if (values.length > dataRowIndex) {
        const dataTypes = [];
        
        for (let c = 0; c < Math.min(columnCount, 6); c++) {
          const value = values[dataRowIndex][c];
          let type = 'text';
          
          if (value !== null && value !== undefined) {
            if (typeof value === 'number' || (typeof value === 'string' && !isNaN(Number(value)))) {
              type = 'number';
            } else if (typeof value === 'string') {
              // Check if date
              const datePattern = /^\d{1,4}[-/]\d{1,2}[-/]\d{1,4}$/;
              if (datePattern.test(value)) {
                type = 'date';
              }
            }
          }
          
          const headerName = hasHeaders && values[0][c] !== null && values[0][c] !== undefined
            ? String(values[0][c])
            : `Column ${c+1}`;
            
          dataTypes.push(`${headerName} (${type})`);
        }
        
        if (dataTypes.length > 0) {
          if (columnCount > 6) {
            formattedData += `- Sample columns: ${dataTypes.slice(0, 3).join(', ')}... and ${columnCount - 6} more\n`;
          } else {
            formattedData += `- Columns: ${dataTypes.join(', ')}\n`;
          }
        }
      }
    }
    
    return formattedData;
  };
  
  /**
   * Process a user message and generate an AI response
   */
  const handleSendMessage = async (message = inputValue) => {
    if (!message || (typeof message === 'string' && !message.trim())) return;
    setSuggestionsVisible(false);
    
    // Generate unique IDs for messages
    let userMessageId = generateId();
    let aiMessageId = generateId();
    
    try {
      // Clear input and set processing state
      setInputValue('');
      setProcessingAction(true);
      
      // Add user message to chat
      addMessage({ 
        id: userMessageId,
        role: 'user', 
        content: message,
        timestamp: new Date().toISOString()
      });
      
      // If using selection, refresh the data first to ensure we have the latest
      if (useSelection && currentSelection && currentSelection.address) {
        // Refresh the selection data to ensure we have the latest
        console.log("Refreshing selection data before processing message");
        await refreshSelectionData();
      }
      
      // If using selection, retrieve the data but don't add it to chat
      if (useSelection && currentSelection && currentSelection.address) {
        // Get the raw values directly from Excel to show the current data
        const directRangeResult = await excelService.getData(currentSelection.address, {
          includeAllRows: true,
          preserveTypes: true
        });
        
        if (directRangeResult && directRangeResult.success && directRangeResult.data) {
          const hasHeaders = directRangeResult.data.length > 1; // Assume headers if more than 1 row
          
          // Save the data for displaying in the dialog/modal
          setSelectionData({
            values: directRangeResult.data,
            hasHeaders,
            address: currentSelection.address,
            rowCount: directRangeResult.rowCount || directRangeResult.data.length,
            columnCount: directRangeResult.columnCount || (directRangeResult.data[0]?.length || 0)
          });
        }
      }
      
      // Add AI "thinking" message with clear range information
      let thinkingMessage = 'AI is processing your request';
      
      // Customize thinking message based on mode
      if (mode.key === 'AGENT') {
        thinkingMessage = 'AI is analyzing your request and preparing to make changes to your spreadsheet';
      }
      
      addMessage({ 
        id: aiMessageId,
        role: 'assistant', 
        content: thinkingMessage, 
        isThinking: true,
        timestamp: new Date().toISOString()
      });
      
      console.log("Starting to process message:", message);
      console.log("Current mode:", mode.key);
      
      // Always force refresh Excel context before processing to ensure we have the latest data
      console.log("Loading Excel context...");
      const excelContextResult = await loadExcelContext(true);
      console.log("Excel context loaded:", excelContextResult ? "success" : "failed");
      setExcelDataLoaded(true);
      
      // Update the current selection if selection mode is active
      if (useSelection && excelContextResult && excelContextResult.selection) {
        // Only update selection if the address has changed
        if (!currentSelection || 
            excelContextResult.selection.address !== currentSelection.address || 
            excelContextResult.selection.rowCount !== currentSelection.rowCount ||
            excelContextResult.selection.columnCount !== currentSelection.columnCount) {
          console.log('Selection changed during message processing');
          setCurrentSelection(excelContextResult.selection);
        }
      }
      
      // Analyze user request to determine intent
      const lowerMessage = message.toLowerCase();
      
      // Extract column information if present
      const columnMatch = lowerMessage.match(/column\s+([a-z])/i);
      let targetColumn = null;
      if (columnMatch && columnMatch[1]) {
        targetColumn = columnMatch[1].toUpperCase();
        console.log(`Column identified: ${targetColumn}`);
      }
      
      // Check for counting questions
      const countingQuestion = 
        lowerMessage.includes('how many') || 
        lowerMessage.includes('count') || 
        lowerMessage.match(/number of/i) ||
        lowerMessage.includes('total');
        
      // Check for specific values to count
      const valuesToCheck = [
        { keyword: 'returning', searchTerm: 'Returning' },
        { keyword: 'new', searchTerm: 'New' },
        { keyword: 'active', searchTerm: 'Active' },
        { keyword: 'inactive', searchTerm: 'Inactive' },
        { keyword: 'pending', searchTerm: 'Pending' },
        { keyword: 'completed', searchTerm: 'Completed' }
      ];
      
      // Build context data to provide to OpenAI - we'll collect data but prioritize OpenAI's analysis
      let enrichedContext = "";
      let dataCollected = false;
      
      // If using selection, collect data
      if (useSelection && currentSelection && currentSelection.address) {
        try {
          console.log("Fetching data for selected range:", currentSelection.address);
          
          // Get the raw values directly from Excel with improved method to ensure ALL rows are captured
          const directRangeResult = await excelService.getData(currentSelection.address, {
            includeAllRows: true,     // Explicitly request all rows
            preserveTypes: true       // Preserve data types for better analysis
          });
          
          if (directRangeResult && directRangeResult.success && directRangeResult.data) {
            console.log("Direct range data retrieved successfully");
            console.log(`Range dimensions: ${directRangeResult.rowCount} rows × ${directRangeResult.columnCount} columns`);
            console.log("Sample of direct range data:", 
                       JSON.stringify(directRangeResult.data.slice(0, Math.min(3, directRangeResult.data.length))));
            
            // Verify we have the expected amount of data
            if (directRangeResult.rowCount === 0 || directRangeResult.columnCount === 0) {
              console.error("Retrieved empty range data");
              updateLastMessage({ 
                id: aiMessageId,
                content: "I couldn't analyze your selection as it appears to be empty. Please select a range with data and try again.",
                isThinking: false,
                isError: true
              });
              setProcessingAction(false);
              return;
            }
            
            // Store direct values for reliable analysis
            const directValues = directRangeResult.data;
            
            // Double-check row count against expected selection
            if (currentSelection.rowCount > 0 && directValues.length < currentSelection.rowCount) {
              console.warn(`Row count mismatch: expected ${currentSelection.rowCount}, got ${directValues.length}`);
            }
            
            // Keep track of the exact range successfully retrieved 
            const retrievedRange = {
              address: directRangeResult.address || currentSelection.address,
              rowCount: directValues.length,
              columnCount: directValues[0] ? directValues[0].length : 0
            };
            
            // Get a fresh copy of the current selection to ensure we have the latest data
            const freshSelectionInfo = await excelService.getSelectedRange();
            
            // Verify selection consistency between operations
            let selectionConsistent = false;
            if (freshSelectionInfo && freshSelectionInfo.success && freshSelectionInfo.data) {
              // Update current selection to latest regardless
                setCurrentSelection({
                  address: freshSelectionInfo.data.address,
                  rowCount: freshSelectionInfo.data.rowCount,
                  columnCount: freshSelectionInfo.data.columnCount,
                  values: freshSelectionInfo.data.values
                });
              
              if (freshSelectionInfo.data.address === retrievedRange.address &&
                  freshSelectionInfo.data.rowCount === retrievedRange.rowCount &&
                  freshSelectionInfo.data.columnCount === retrievedRange.columnCount) {
                console.log("Fresh selection data verified and consistent with direct data");
                selectionConsistent = true;
              } else {
                console.warn("Selection may have changed between operations:", 
                  `Retrieved: ${retrievedRange.address} (${retrievedRange.rowCount}×${retrievedRange.columnCount})`,
                  `Current: ${freshSelectionInfo.data.address} (${freshSelectionInfo.data.rowCount}×${freshSelectionInfo.data.columnCount})`);
              }
            }
            
            // Use the directly retrieved values for analysis to guarantee complete data
            const useValues = directValues;
            
            // Get detailed metadata about the selection to enhance analysis
            const selectionData = await getSelectionData(retrievedRange.address, { 
              forceIncludeAllRows: true,  // Make sure we include all rows in analysis
              detectHeaders: true,        // Try to auto-detect headers
              includeFormulas: true,      // Include original formulas if present
              preserveDataTypes: true     // Keep original data types
            });
            
            // Log the analysis results
            console.log("Selection analysis complete:", 
                      selectionData.success ? "Success" : "Failed", 
                      selectionData.error || "");
            
            // Build the analysis context only if we have valid data
            if (selectionData && selectionData.success && selectionData.data) {
              dataCollected = true;
              
              // Log the values for comparison to ensure they match
              console.log("Comparing data sources lengths:", 
                         "Direct:", useValues.length, "×", useValues[0]?.length || 0,
                         "Analyzed:", selectionData.data.values.length, "×", selectionData.data.values[0]?.length || 0);
              
              // Determine if headers exist in the data
              // Start by using the value from selection data (which does analysis)
              let hasHeaders = selectionData.data.hasHeaders;
              
              // Additional header detection if not already detected
              if (!hasHeaders && useValues.length > 1) {
                // Look for patterns indicating headers (first row different format from others)
                const firstRow = useValues[0];
                let headerLikelihood = 0;
                
                // Check if first row has text while other rows have numbers
                let firstRowTextCount = 0;
                let otherRowsTextCount = 0;
                let otherRowsCount = 0;
                
                // Count text cells in first row
                for (let c = 0; c < firstRow.length; c++) {
                  if (typeof firstRow[c] === 'string' && !isNumeric(firstRow[c])) {
                    firstRowTextCount++;
                  }
                }
                
                // Count text cells in other rows
                for (let r = 1; r < Math.min(useValues.length, 6); r++) {
                  const row = useValues[r];
                  for (let c = 0; c < row.length; c++) {
                    if (typeof row[c] === 'string' && !isNumeric(row[c])) {
                      otherRowsTextCount++;
                    }
                  }
                  otherRowsCount++;
                }
                
                const firstRowTextRatio = firstRowTextCount / firstRow.length;
                const otherRowsTextRatio = otherRowsTextCount / (otherRowsCount * firstRow.length);
                
                // If first row has significantly more text than other rows, likely headers
                if (firstRowTextRatio > 0.5 && firstRowTextRatio > otherRowsTextRatio * 1.5) {
                  hasHeaders = true;
                }
                
                // Check if first row has shorter text (often label/header) while other rows have longer content
                let firstRowAvgLength = 0;
                let otherRowsAvgLength = 0;
                
                // Calculate average text length in first row
                for (let c = 0; c < firstRow.length; c++) {
                  if (typeof firstRow[c] === 'string') {
                    firstRowAvgLength += firstRow[c].length;
                  }
                }
                firstRowAvgLength = firstRowAvgLength / firstRow.length;
                
                // Calculate average text length in other rows
                let totalLength = 0;
                let cellCount = 0;
                for (let r = 1; r < Math.min(useValues.length, 6); r++) {
                  const row = useValues[r];
                  for (let c = 0; c < row.length; c++) {
                    if (typeof row[c] === 'string') {
                      totalLength += row[c].length;
                      cellCount++;
                    }
                  }
                }
                otherRowsAvgLength = cellCount > 0 ? totalLength / cellCount : 0;
                
                // If other rows have significantly longer text than first row, likely headers
                if (otherRowsAvgLength > firstRowAvgLength * 1.5) {
                  hasHeaders = true;
                }
              }
              
              // Start with basic selection information
              enrichedContext = `\n\nI'm analyzing the Excel data you've selected in range ${retrievedRange.address}.\n`;
              enrichedContext += `The selection contains ${useValues.length} rows and ${useValues[0]?.length || 0} columns.\n`;
              
              // Add header information if detected
              if (hasHeaders) {
                enrichedContext += "\nThe first row contains column headers.\n";
              }
              
              // Format the data in a plain tabular format that's easier to understand
              enrichedContext += "\nHere's the exact data in your selection:\n\n";
              
              // Helper function to check if a value looks numeric
              function isNumeric(val) {
                return !isNaN(parseFloat(val)) && isFinite(val);
              }
              
              // Helper function to format a cell value for display
              function formatCellValue(val) {
                if (val === null || val === undefined) return '';
                
                // For numbers, keep precision but format attractively
                if (typeof val === 'number') {
                  // Format integer without decimals, float with up to 4 decimal places
                  return Number.isInteger(val) ? val.toString() : val.toFixed(4).replace(/\.?0+$/, '');
                }
                
                // For dates, format consistently
                if (val instanceof Date) {
                  return val.toISOString().split('T')[0];
                }
                
                // For strings, escape any special characters that might break formatting
                return String(val);
              }
              
              if (useValues && useValues.length > 0) {
                // Remove the other formats and only keep the JSON representation
                
                // Convert to JSON for more accurate analysis
                enrichedContext += "\n\n=== COMPLETE DATA IN JSON FORMAT ===\n";
                enrichedContext += "This JSON array contains the complete dataset with preserved data types. Each object represents one row with properly named columns.\n";
                
                try {
                  // Extract headers from first row or generate generic ones
                  const headers = hasHeaders && useValues.length > 0 ? 
                    useValues[0].map((header, idx) => {
                      // Clean headers to ensure they're valid JSON object keys
                      let cleanHeader = header ? String(header).trim() : `Column_${idx+1}`;
                      
                      // Replace spaces and special characters to ensure valid JSON property names
                      cleanHeader = cleanHeader.replace(/[^\w\d_]/g, '_');
                      
                      // Ensure it starts with a letter or underscore (not a number)
                      if (/^\d/.test(cleanHeader)) {
                        cleanHeader = 'Col_' + cleanHeader;
                      }
                      
                      // Ensure uniqueness in case of duplicates
                      return cleanHeader;
                    }) : 
                    Array.from({length: useValues[0].length}, (_, i) => `Column_${i+1}`);
                  
                  // Convert the data to JSON objects with proper headers as keys
                  const jsonData = [];
                  
                  // Start from row 1 if there are headers, otherwise start from row 0
                  const jsonStartRow = hasHeaders ? 1 : 0;
                  
                  // Ensure we iterate through ALL rows
                  for (let r = jsonStartRow; r < useValues.length; r++) {
                    const rowObject = {};
                    for (let c = 0; c < useValues[r].length; c++) {
                      const header = headers[c];
                      // Preserve the original data type when possible
                      let cellValue = useValues[r][c];
                      
                      // Skip null/undefined values
                      if (cellValue === null || cellValue === undefined) {
                        rowObject[header] = null;
                        continue;
                      }
                      
                      // Try to convert numeric strings to actual numbers for better analysis
                      if (typeof cellValue === 'string') {
                        // Check if it's a date string
                        const dateValue = new Date(cellValue);
                        if (!isNaN(dateValue.getTime()) && 
                            (cellValue.includes('-') || cellValue.includes('/') || cellValue.includes(','))) {
                          cellValue = dateValue;
                        } else {
                          // Check if it's a numeric string
                        const numericValue = Number(cellValue);
                        if (!isNaN(numericValue) && cellValue.trim() !== '') {
                          cellValue = numericValue;
                          }
                          
                          // Check if it's a boolean
                          if (cellValue.toLowerCase() === 'true') cellValue = true;
                          if (cellValue.toLowerCase() === 'false') cellValue = false;
                        }
                      }
                      
                      rowObject[header] = cellValue;
                    }
                    jsonData.push(rowObject);
                  }
                  
                  // Verify we've created the expected number of JSON objects
                  console.log(`JSON data conversion complete: ${jsonData.length} rows created from ${useValues.length - jsonStartRow} source rows`);
                  
                  // If rows are missing, log a warning
                  if (jsonData.length < useValues.length - jsonStartRow) {
                    console.warn(`Missing rows in JSON conversion: expected ${useValues.length - jsonStartRow}, got ${jsonData.length}`);
                  }
                  
                  // Perform schema validation on the JSON data
                  const validationResult = validateDataSchema(jsonData, headers);
                  
                  // Pre-calculate important metrics and counts for common analyses
                  const preCalculatedMetrics = calculateDataMetrics(jsonData, headers);
                  
                  // We'll capture the exact JSON data that's sent to OpenAI at the prompt stage instead
                  // This ensures we capture exactly what the AI sees
                  
                  // Add the JSON data to the context with pretty formatting and explicit code block
                  enrichedContext += "```json\n";
                    enrichedContext += JSON.stringify(jsonData, null, 2);
                  enrichedContext += "\n```\n";
                    
                  // Add data dimensions verification to confirm full dataset inclusion
                  enrichedContext += `\n!!! VERIFICATION: This JSON data contains ALL ${jsonData.length} rows from the selection (rows ${jsonStartRow+1} to ${useValues.length}), with preserved data types !!!\n`;
                  
                  // Add structured analysis information
                  if (validationResult.valid) {
                    // Add schema validation confirmation for the AI
                    enrichedContext += "\n\nData schema validation: Passed ✓";
                    
                    // Add schema information for better AI understanding
                    enrichedContext += "\nSchema structure:";
                    enrichedContext += `\n- Rows: ${jsonData.length}`;
                    enrichedContext += `\n- Columns: ${headers.length}`;
                    enrichedContext += `\n- Column names: ${headers.join(', ')}`;
                    enrichedContext += `\n- Column types: ${JSON.stringify(validationResult.columnTypes)}`;
                  } else {
                    // Log schema validation issues but still include the data
                    console.warn("Schema validation warnings:", validationResult.issues);
                    enrichedContext += "\n\nData schema validation: Warnings ⚠️";
                    enrichedContext += "\nThe following issues were detected:";
                    validationResult.issues.forEach(issue => {
                      enrichedContext += `\n- ${issue}`;
                    });
                  }
                    
                    // Add pre-calculated metrics for more reliable analysis
                    enrichedContext += "\n\nPre-calculated data metrics:";
                  
                  // Add data completeness information
                  if (validationResult.completeness !== undefined) {
                    enrichedContext += `\n- Data completeness: ${Math.round(validationResult.completeness * 100)}%`;
                  }
                  
                  // Add row count confirmation
                  enrichedContext += `\n- Total rows analyzed: ${jsonData.length}`;
                  enrichedContext += `\n- Original Excel range: ${retrievedRange.address} (${retrievedRange.rowCount} rows × ${retrievedRange.columnCount} columns)`;
                    
                    // Add special status verification section if available
                    if (preCalculatedMetrics.statusVerification) {
                      enrichedContext += preCalculatedMetrics.statusVerification;
                    }
                    
                    // Add verification data for critical data points
                    if (preCalculatedMetrics.verificationData) {
                      enrichedContext += "\n\nROW-BY-ROW VERIFICATION DATA:";
                      
                      // Add sample of the exact row data for verification
                      const sampleSize = Math.min(jsonData.length, 10);
                      for (let i = 0; i < sampleSize; i++) {
                        enrichedContext += `\nRow ${i+1} contains: `;
                        const rowData = preCalculatedMetrics.verificationData.rowByRowValidation[i];
                        if (rowData) {
                          const rowDataStr = Object.entries(rowData)
                            .map(([key, val]) => `${key}: "${val}"`)
                            .join(", ");
                          enrichedContext += rowDataStr;
                        }
                      }
                  
                    // Add last row explicitly to verify it's included
                    if (jsonData.length > 10) {
                      const lastRowIdx = jsonData.length - 1;
                      enrichedContext += `\n\nLast row (Row ${lastRowIdx+1}) contains: `;
                      const lastRowData = preCalculatedMetrics.verificationData.rowByRowValidation[lastRowIdx];
                      if (lastRowData) {
                        const rowDataStr = Object.entries(lastRowData)
                            .map(([key, val]) => `${key}: "${val}"`)
                            .join(", ");
                          enrichedContext += rowDataStr;
                        }
                      }
                      
                      // Add additional verification data for specific columns
                      Object.entries(preCalculatedMetrics.verificationData)
                        .filter(([key]) => key !== 'rowByRowValidation' && key !== 'totalRows')
                        .forEach(([column, data]) => {
                          if (data.countBreakdown) {
                            enrichedContext += `\n\nVerification for column "${column}":\n`;
                            enrichedContext += data.countBreakdown.join("\n");
                          }
                        });
                  }
                  
                  // Store JSON data for use if in agent mode
                  if (mode.key === 'AGENT') {
                    setOpenAIJsonData(jsonData);
                  }
                } catch (error) {
                  console.error("Error converting data to JSON:", error);
                  enrichedContext += `Error converting data to JSON format: ${error.message}. Using tabular format only.`;
                }
              }
              
              // Add basic statistics about the data
              enrichedContext += `\n\nThis selection contains ${useValues.length} rows and ${useValues[0]?.length || 0} columns.`;
              enrichedContext += `\nNon-empty cells: ${selectionData.data.summary.nonEmptyCells}`;
              enrichedContext += `\nEmpty cells: ${selectionData.data.summary.emptyCells}`;
              
              // Add value frequencies if available
              // ... rest of existing code for value frequencies
              
              // Log what we're sending to OpenAI
              console.log("Data context prepared for OpenAI");
              console.log(`Context length: ${enrichedContext.length} characters`);
              console.log(`Data rows included: ${useValues.length}`);
            }
          } else {
            console.error("Failed to get direct range data:", directRangeResult?.error || "Unknown error");
            updateLastMessage({ 
              id: aiMessageId,
              content: "I couldn't retrieve your selected data. Please try selecting the range again.",
              isThinking: false,
              isError: true
            });
            setProcessingAction(false);
            return;
          }
        } catch (error) {
          console.error("Error collecting enrichment data:", error);
          updateLastMessage({
            id: aiMessageId,
            content: `I encountered an error analyzing your data: ${error.message}. Please try again with a different selection.`,
            isThinking: false,
            isError: true
          });
          setProcessingAction(false);
          return;
        }
      }
      
      // Process based on AI mode
      if (mode.key === 'AGENT') {
        // Agent Mode - directly apply changes to the spreadsheet
        console.log("Processing in Agent Mode - will apply changes to spreadsheet");
        
        // First, check if we have selected data to work with
        if (!useSelection || !currentSelection || !currentSelection.address) {
          console.warn("Agent Mode requires selection, but no selection found");
          updateLastMessage({
            id: aiMessageId,
            content: "To use Agent Mode, please select a range of data first. This helps me understand exactly which data to modify.",
            isThinking: false
          });
          setProcessingAction(false);
          return;
        }
        
        try {
          // Create a simple string prompt that includes all necessary context
          const completePrompt = `
I need Office.js code to modify an Excel spreadsheet based on the following request:

USER REQUEST: ${message}

EXCEL CONTEXT:
- Selected Range: ${currentSelection.address}
- Range Dimensions: ${currentSelection.rowCount} rows × ${currentSelection.columnCount} columns
${excelContextResult?.activeWorksheet ? `- Active Worksheet: ${excelContextResult.activeWorksheet}` : ''}

IMPORTANT REQUIREMENTS:
1. ONLY work with the selected data range: ${currentSelection.address}
2. Generate EXECUTABLE Office.js code that implements the user's request
3. Your code MUST be wrapped in Excel.run() and include proper context.sync() calls
4. Include error handling with try/catch in your code
5. DO NOT reference cells outside the selected range
6. Focus on implementing ONE clear modification per request
7. IMPORTANT: First explain what your code will do, then provide the code

The code should follow this structure:
\`\`\`javascript
await Excel.run(async (context) => {
  try {
    // Your implementation code here
    // ...
    await context.sync();
    return { success: true, message: "Description of what was completed" };
  } catch (error) {
    console.error("Error:", error);
    return { success: false, error: error.message };
  }
});
\`\`\`

Please implement this change using best practices for Office.js:
- Use sheet.getRange() with explicit addresses
- Use Excel.NumberFormat enums for formatting when available
- Include proper await context.sync() calls after loading properties
- For conditional formatting, use the proper Excel.ConditionalFormatType enums
`;

          console.log("Sending simple string prompt for Agent Mode");
          
          // Generate the Office.js code using a simple string prompt
          const openAIResponse = await generateText(completePrompt);

          if (!openAIResponse.success) {
            console.error("Error generating Office.js code:", openAIResponse.error);
            updateLastMessage({
              id: aiMessageId,
              content: `I encountered an error while generating the code for your request: ${openAIResponse.error}. Please try again with a more specific request.`,
              isThinking: false,
              isError: true
            });
            setProcessingAction(false);
            return;
          }

          // Extract the code from the response
          const responseContent = openAIResponse.content;
          const codeMatch = responseContent.match(/```javascript([\s\S]*?)```/);
          let codeToExecute = null;

          if (codeMatch && codeMatch[1]) {
            codeToExecute = codeMatch[1].trim();
          } else {
            // If no code block found, check if the entire response might be code
            if (responseContent.includes('Excel.run(') && responseContent.includes('context.sync()')) {
              codeToExecute = responseContent.trim();
            }
          }

          if (!codeToExecute) {
            console.error("No executable code found in the response");
            updateLastMessage({
              id: aiMessageId,
              content: `I understood your request, but couldn't generate proper Excel code to implement it. Please try a simpler request or provide more details about what you need.`,
              isThinking: false
            });
            setProcessingAction(false);
            return;
          }

          // Format the message to show to the user before execution
          let executionMessage = responseContent.replace(/```javascript([\s\S]*?)```/, '```javascript\n// Code will be executed automatically\n```');
          
          // Update the message to show what will be done
          updateLastMessage({
            id: aiMessageId,
            content: executionMessage,
            isThinking: false
          });

          // Execute the code after a short delay to allow the user to see what's happening
          setTimeout(async () => {
            try {
              console.log("Executing Office.js code:", codeToExecute);
              
              // Execute the code
              const executionResult = await excelService.execute(codeToExecute);
              
              if (executionResult.success) {
                console.log("Code execution successful:", executionResult);
                
                // Add a success message
                addMessage({
                  id: generateId(),
                  role: 'assistant',
                  content: `✅ Changes applied successfully to your spreadsheet.${executionResult.message ? ' ' + executionResult.message : ''}`,
                  timestamp: new Date().toISOString()
                });
                
                // Refresh the selection data to show the updated values
                await refreshSelectionData();
              } else {
                console.error("Code execution failed:", executionResult.error);
                
                // Add an error message
                addMessage({
                  id: generateId(),
                  role: 'assistant',
                  content: `❌ Error applying changes: ${executionResult.error}. Please try again with a more specific request.`,
                  isError: true,
                  timestamp: new Date().toISOString()
                });
              }
              
              setProcessingAction(false);
            } catch (error) {
              console.error("Error during code execution:", error);
              
              // Add an error message
              addMessage({
                id: generateId(),
                role: 'assistant',
                content: `❌ An unexpected error occurred: ${error.message}. Please try again with a different request.`,
                isError: true,
                timestamp: new Date().toISOString()
              });
              
              setProcessingAction(false);
            }
          }, 1500);
        } catch (error) {
          console.error("Error processing in Agent Mode:", error);
          updateLastMessage({
            id: aiMessageId,
            content: `I'm sorry, but an error occurred while processing your request: ${error.message}. Please try again.`,
            isThinking: false,
            isError: true
          });
          setProcessingAction(false);
        }
        
        return;
      }
      
      // For Ask and Prompt modes, continue with the original logic...
      // ... (existing code continues)
    } catch (error) {
      console.error("Error in chat processing:", error);
      updateLastMessage({
        id: aiMessageId,
        content: `I'm sorry, but an error occurred: ${error.message || "Unknown error"}`,
        isThinking: false,
        isError: true
      });
      setProcessingAction(false);
    }
  };
  
  /**
   * Handle undoing the last Excel operation
   */
  const handleUndo = async () => {
    try {
      setProcessingAction(true);
      const result = await undoLastOperation();
      
      if (result.success) {
        addMessage({ 
          role: 'assistant',
          content: 'I\'ve undone the last operation. ' + 
                   (result.message || 'Your spreadsheet has been restored to its previous state.'),
          isSuccess: true
        });
      } else {
        addMessage({ 
          role: 'assistant',
          content: 'I couldn\'t undo the last operation: ' + 
                   (result.error || 'Unknown error'),
          isError: true
        });
      }
    } catch (err) {
      console.error('Error undoing operation:', err);
      addMessage({ 
          role: 'assistant',
        content: `Error undoing operation: ${err.message || 'Unknown error occurred'}`,
          isError: true
        });
    } finally {
      setProcessingAction(false);
    }
  };
  
  /**
   * Handle input change
   */
  const handleInputChange = (event, newValue) => {
    try {
      // Ensure we're always working with a string value
      const textValue = typeof newValue === 'string' ? newValue : '';
      setInputValue(textValue);
      } catch (error) {
      console.error("Error in handleInputChange:", error);
      setInputValue('');
    }
  };
  
  /**
   * Handle key down events in the input
   */
  const handleKeyDown = (event) => {
    try {
      if (event && event.key === 'Enter' && !event.shiftKey) {
        event.preventDefault();
        handleSendMessage();
      }
    } catch (error) {
      console.error("Error in handleKeyDown:", error);
    }
  };
  
  /**
   * Handle clicking a suggestion
   */
  const handleSuggestionClick = (suggestion) => {
    setInputValue(suggestion);
    setSuggestionsVisible(false);
    
    // Auto-send the suggestion after a short delay
    setTimeout(() => {
      handleSendMessage(suggestion);
    }, 100);
  };
  
  /**
   * Handle changing the AI mode
   */
  const handleModeChange = (newMode) => {
    setMode(newMode);
    setSuggestionsVisible(true);
  };
  
  /**
   * Render the chat content based on state
   */
  const renderChatContent = () => {
    // Show API key error if no key is set
    if (!apiKey) {
      return (
        <div className={classes.messagesContainer}>
          <MessageBar intent="error">
            <MessageBarBody>
              Please set your OpenAI API key in the settings to use the AI assistant.
            </MessageBarBody>
          </MessageBar>
        </div>
      );
    }
    
    // Show empty state if no messages
    if (messages.length === 0) {
      return (
        <div className={classes.messagesContainer}>
          <EmptyChat 
            mode={mode}
            suggestions={SUGGESTIONS[mode.key]}
            onSuggestionClick={handleSuggestionClick}
          />
        </div>
      );
    }
    
    // Show messages and suggestions
    return (
      <div className={classes.messagesContainer} ref={chatContainerRef}>
        {messages.map((message, index) => (
          <ChatMessage
            key={index}
            message={message}
            style={{ '--message-index': index }}
            onResend={() => {
              if (message.role === 'user') {
                handleSendMessage(message.content);
              }
            }}
          />
        ))}
        <div ref={messagesEndRef} />
        
        {suggestionsVisible && messages.length > 0 && !isProcessing && !processingAction && (
          <SuggestionsList 
            suggestions={SUGGESTIONS[mode.key]} 
            onSuggestionClick={handleSuggestionClick}
          />
        )}
      </div>
    );
  };

  // Generate a simple ID for messages
  function generateId() {
    return Math.random().toString(36).substring(2, 15) + Math.random().toString(36).substring(2, 15);
  }

  /**
   * Validates the schema of JSON data converted from Excel
   * @param {Array} jsonData - Array of row objects
   * @param {Array} headers - Array of column headers
   * @returns {Object} Validation result with status and issues
   */
  function validateDataSchema(jsonData, headers) {
    // Initialize validation result
    const result = {
      valid: true,
      issues: [],
      insights: [],
      columnTypes: {}
    };
    
    // Check if we have data
    if (!jsonData || !Array.isArray(jsonData) || jsonData.length === 0) {
      result.valid = false;
      result.issues.push("No data rows found");
      return result;
    }
    
    // Check if headers are defined
    if (!headers || !Array.isArray(headers) || headers.length === 0) {
      result.valid = false;
      result.issues.push("No column headers defined");
      return result;
    }
    
    // Analyze column types with improved type detection
    const columnDataTypes = {};
    const columnTotals = {};
    const columnNonNull = {};
    const uniqueValues = {};
    const columnMin = {};
    const columnMax = {};
    
    // Initialize tracking objects
    headers.forEach(header => {
      columnDataTypes[header] = new Set();
      columnTotals[header] = 0;
      columnNonNull[header] = 0;
      uniqueValues[header] = new Set();
      columnMin[header] = Number.MAX_VALUE;
      columnMax[header] = Number.MIN_VALUE;
    });
    
    // Track row completeness
    let completeRows = 0;
    let incompleteRows = 0;
    
    // Process each row with improved data validation
    jsonData.forEach((row, rowIndex) => {
      let isRowComplete = true;
      let headerCount = 0;
      
      // Analyze each header/column
      headers.forEach(header => {
        // Check if the row has this header
        if (Object.prototype.hasOwnProperty.call(row, header)) {
          headerCount++;
          const value = row[header];
          
          // Track data types
          if (value === null || value === undefined) {
            // This column has a null value in this row
            isRowComplete = false;
          } else {
            columnNonNull[header]++;
            const valueType = typeof value;
            columnDataTypes[header].add(valueType);
            uniqueValues[header].add(String(value));
            
            // Sum numeric values and track min/max
            if (valueType === 'number') {
              columnTotals[header] += value;
              columnMin[header] = Math.min(columnMin[header], value);
              columnMax[header] = Math.max(columnMax[header], value);
            } else if (valueType === 'string') {
              // Try to convert strings to numbers if they look numeric
              const numValue = Number(value);
              if (!isNaN(numValue)) {
                // Track as potential numeric even if stored as string
                columnTotals[header] += numValue;
                columnMin[header] = Math.min(columnMin[header], numValue);
                columnMax[header] = Math.max(columnMax[header], numValue);
              }
            }
          }
        } else {
          // This row is missing this header
          isRowComplete = false;
          result.issues.push(`Row ${rowIndex+1} is missing column "${header}"`);
        }
      });
      
      // Check for extra properties not in headers
      const extraProps = Object.keys(row).filter(prop => !headers.includes(prop));
      if (extraProps.length > 0) {
        result.issues.push(`Row ${rowIndex+1} has extra properties: ${extraProps.join(', ')}`);
      }
      
      // Track completeness
      if (isRowComplete && headerCount === headers.length) {
        completeRows++;
      } else {
        incompleteRows++;
      }
    });
    
    // Determine primary column types with improved classification
    headers.forEach(header => {
      const types = Array.from(columnDataTypes[header]);
      
      // Determine type with more sophisticated logic
      if (types.length === 0) {
        result.columnTypes[header] = 'empty';
      } else if (types.length === 1) {
        if (types[0] === 'number') {
          result.columnTypes[header] = 'number';
        } else if (types[0] === 'string') {
          // Check if the strings are actually dates or numbers
          const sampleValues = Array.from(uniqueValues[header]).slice(0, 5);
          
          // Check for date strings
          const possibleDates = sampleValues.map(v => new Date(v)).filter(d => !isNaN(d.getTime()));
          if (possibleDates.length === sampleValues.length) {
            result.columnTypes[header] = 'date';
          } else {
            // Check for numeric strings
            const possibleNumbers = sampleValues.map(v => Number(v)).filter(n => !isNaN(n));
            if (possibleNumbers.length === sampleValues.length) {
              result.columnTypes[header] = 'numeric string';
            } else {
              result.columnTypes[header] = 'string';
            }
          }
        } else if (types[0] === 'boolean') {
          result.columnTypes[header] = 'boolean';
        } else {
        result.columnTypes[header] = types[0];
        }
      } else if (types.includes('number') && types.includes('string')) {
        // Check if strings might actually be numbers or dates
        result.columnTypes[header] = 'mixed (number/string)';
      } else {
        result.columnTypes[header] = 'mixed';
      }
      
      // Fix min/max values for empty columns
      if (columnMin[header] === Number.MAX_VALUE) columnMin[header] = null;
      if (columnMax[header] === Number.MIN_VALUE) columnMax[header] = null;
    });
    
    // Calculate and add completeness ratio
    const completenessRatio = jsonData.length > 0 ? completeRows / jsonData.length : 0;
    result.completeness = completenessRatio;
    
    if (completenessRatio < 1) {
      if (completenessRatio < 0.5) {
        result.valid = false; // Mark as invalid if more than half rows are incomplete
        result.issues.push(`Data is highly incomplete (only ${Math.round(completenessRatio*100)}% complete rows)`);
      } else {
        result.issues.push(`${incompleteRows} rows (${Math.round((1-completenessRatio)*100)}%) have missing values`);
      }
    } else {
      result.insights.push("All rows have complete data");
    }
    
    // Generate insights about potential correlations and data summaries
    const numericColumns = headers.filter(h => 
      result.columnTypes[h] === 'number' || 
      result.columnTypes[h] === 'numeric string' ||
      result.columnTypes[h] === 'mixed (number/string)');
    
    if (numericColumns.length > 0) {
      result.insights.push(`Numeric columns available for analysis: ${numericColumns.join(', ')}`);
      
      // Add basic column statistics for numeric columns
      numericColumns.forEach(col => {
        const nonNullCount = columnNonNull[col];
        if (nonNullCount > 0) {
          const avg = columnTotals[col] / nonNullCount;
          const min = columnMin[col];
          const max = columnMax[col];
          result.insights.push(`Column "${col}" has average value of ${avg.toFixed(2)}, ranging from ${min} to ${max}`);
        }
      });
    }
    
    return result;
  }

  /**
   * Calculate detailed metrics from the data for reliable analysis
   * @param {Array} jsonData - The JSON data array
   * @param {Array} headers - The column headers
   * @returns {Object} Detailed metrics and counts
   */
  function calculateDataMetrics(jsonData, headers) {
    // Initialize the metrics object with more comprehensive structure
    const metrics = {
      valueCounts: {},         // Count of each value by column
      valueLocations: {},      // Row indices where each value appears
      numericStats: {},        // Numeric statistics by column
      verificationData: {},    // Verification data with exact rows
      statusVerification: "",  // Status verification text
      totalSummary: {},        // Summary of entire dataset
      correlations: {}         // Potential correlations between columns
    };
    
    // Skip processing if no data
    if (!jsonData || !Array.isArray(jsonData) || jsonData.length === 0 || !headers) {
      return metrics;
    }
    
    // Add dataset summary for verification
    metrics.totalSummary = {
      totalRows: jsonData.length,
      totalColumns: headers.length,
      columnNames: headers
    };
    
    // Add verification data with exact row listings for key values
    metrics.verificationData.totalRows = jsonData.length;
    metrics.verificationData.rowByRowValidation = {};
    
    // Process each row for verification with improved data tracking
    jsonData.forEach((row, rowIndex) => {
      // Store simplified data for each row to easily validate specific rows
      const rowData = {};
      Object.entries(row).forEach(([key, value]) => {
        // Convert to string for consistent comparison
        rowData[key] = value === null || value === undefined ? '' : String(value);
      });
      
      // Store indexed by both number and 1-based position (humans use 1-indexed rows)
      metrics.verificationData.rowByRowValidation[rowIndex] = rowData;
      metrics.verificationData.rowByRowValidation[`Row ${rowIndex + 1}`] = rowData;
    });
    
    // Initialize structures for each column
    headers.forEach(header => {
      metrics.valueCounts[header] = {};
      metrics.valueLocations[header] = {};
      
      // First pass: identify data types, count occurrences, and collect statistics
      let numericValues = [];
      let allNumeric = true;
      let isStatusColumn = header.toLowerCase().includes('status');
      
      jsonData.forEach((row, rowIndex) => {
        const value = row[header];
        
        // Skip null values
        if (value === null || value === undefined) {
          return;
        }
        
        // Convert to string for consistent counting
        const strValue = String(value);
        
        // Check if it's a number for statistics
        if (typeof value === 'number') {
          numericValues.push(value);
        } else if (typeof value === 'string') {
          const numValue = Number(value);
          if (!isNaN(numValue)) {
            numericValues.push(numValue);
          } else {
            allNumeric = false;
          }
        } else {
          allNumeric = false;
        }
        
        // Count occurrences with improved tracking
        metrics.valueCounts[header][strValue] = (metrics.valueCounts[header][strValue] || 0) + 1;
        
        // Track row locations for each value (for complete lists)
        if (!metrics.valueLocations[header][strValue]) {
          metrics.valueLocations[header][strValue] = [];
        }
        // Store 1-indexed row numbers for human readability
        metrics.valueLocations[header][strValue].push(rowIndex + 1);
      });
      
      // Calculate numeric stats with verification checks
      if (numericValues.length > 0) {
        const sum = numericValues.reduce((sum, val) => sum + val, 0);
        const average = sum / numericValues.length;
        
        // Add verification data by showing the exact calculation
        const verificationInfo = {
          values: numericValues.slice(0, 20), // First 20 values for verification
          sumCalculation: numericValues.length <= 10 ? 
            numericValues.join(' + ') + ' = ' + sum : 
            'Sum of ' + numericValues.length + ' values = ' + sum,
          averageCalculation: 'Sum (' + sum + ') ÷ Count (' + numericValues.length + ') = ' + average
        };
        
        metrics.numericStats[header] = {
          count: numericValues.length,
          sum: sum,
          average: average,
          min: Math.min(...numericValues),
          max: Math.max(...numericValues),
          verification: verificationInfo // Include verification data
        };
      }
      
      // Enhanced special handling for Status columns
      if (isStatusColumn) {
        const statusValues = Object.keys(metrics.valueCounts[header]);
        if (statusValues.length > 0 && statusValues.length <= 10) {
          // Create detailed status verification section
          let statusVerification = "\n\nSTATUS VERIFICATION COUNTS:";
          statusValues.forEach(status => {
            const count = metrics.valueCounts[header][status];
            const rowsList = metrics.valueLocations[header][status];
            statusVerification += `\n- "${status}": ${count} occurrences`;
            
            // Add row details for smaller datasets (up to 50 rows total)
            if (jsonData.length <= 50) {
              statusVerification += ` in rows: ${rowsList.join(', ')}`;
            } else if (rowsList.length <= 10) {
              // For larger datasets, only show row numbers for less frequent statuses
              statusVerification += ` in rows: ${rowsList.join(', ')}`;
            }
          });
          metrics.statusVerification = statusVerification;
        }
      }
      
      // Find all unique values for categorical data verification
      if (!allNumeric && Object.keys(metrics.valueCounts[header]).length > 0) {
        const uniqueValues = Object.keys(metrics.valueCounts[header]);
        
        // Add data for verification with improved organization
        if (uniqueValues.length <= 15) { // Only for reasonable number of values
          metrics.verificationData[header] = {
            uniqueValues: uniqueValues,
            countBreakdown: uniqueValues.map(val => {
              const count = metrics.valueCounts[header][val];
              const locations = metrics.valueLocations[header][val];
              return `"${val}": ${count} occurrences in rows ${locations.join(', ')}`;
            })
          };
        }
      }
    });
    
    // Look for correlations between columns (for small datasets)
    if (jsonData.length <= 100 && headers.length >= 2) {
      metrics.correlations = findColumnCorrelations(jsonData, headers);
    }
    
    return metrics;
  }

  /**
   * Find potential correlations between columns
   */
  function findColumnCorrelations(jsonData, headers) {
    const correlations = {};
    
    // For each pair of columns
    for (let i = 0; i < headers.length; i++) {
        for (let j = i + 1; j < headers.length; j++) {
            const col1 = headers[i];
            const col2 = headers[j];
            
            // Track value pairs
            const valuePairs = {};
            
            // Analyze rows
            jsonData.forEach(row => {
                const val1 = row[col1];
                const val2 = row[col2];
                
                // Skip if either value is null
                if (val1 === null || val1 === undefined || val2 === null || val2 === undefined) {
                    return;
                }
                
                // Convert to strings for consistent keys
                const key = `${val1}:${val2}`;
                valuePairs[key] = (valuePairs[key] || 0) + 1;
            });
            
            // Check if there's a strong correlation
            const uniquePairs = Object.keys(valuePairs).length;
            const uniqueFirstValues = new Set(Object.keys(valuePairs).map(k => k.split(':')[0])).size;
            
            // If there's a strong correlation (few unique pairs relative to data size)
            if (uniquePairs <= uniqueFirstValues * 1.5 && uniquePairs <= jsonData.length * 0.5) {
                correlations[`${col1} → ${col2}`] = {
                    pairs: valuePairs,
                    strength: 1 - (uniquePairs / jsonData.length)
                };
            }
        }
    }
    
    return correlations;
  }

  // Add a useEffect to reset OpenAI data when selection changes
  useEffect(() => {
    // Clear the openAIJsonData whenever the selection changes
    if (currentSelection) {
      console.log("Selection changed, clearing previous OpenAI data");
      setOpenAIJsonData(null);
    }
  }, [currentSelection?.address]); // Only trigger when the selection address changes

  // Function to refresh the selection data from Excel
  const refreshSelectionData = async () => {
    if (!useSelection || !currentSelection || !currentSelection.address) {
      console.log("No active selection to refresh");
      return;
    }
    
    try {
      console.log("Refreshing selection data from Excel");
      const directRangeResult = await excelService.getData(currentSelection.address, {
        includeAllRows: true,
        preserveTypes: true
      });
      
      if (directRangeResult && directRangeResult.success && directRangeResult.data) {
        const hasHeaders = directRangeResult.data.length > 1; 
        
        // Update the selection data with fresh data
        setSelectionData({
          values: directRangeResult.data,
          hasHeaders,
          address: currentSelection.address,
          rowCount: directRangeResult.rowCount || directRangeResult.data.length,
          columnCount: directRangeResult.columnCount || (directRangeResult.data[0]?.length || 0),
          refreshTimestamp: new Date().toISOString() // Add timestamp to track when it was refreshed
        });
        
        // Clear any existing OpenAI data to ensure it will be regenerated
        setOpenAIJsonData(null);
        
        console.log("Selection data refreshed successfully");
        return true;
      } else {
        console.error("Failed to refresh range data:", directRangeResult?.error || "Unknown error");
        return false;
      }
    } catch (error) {
      console.error("Error refreshing selection data:", error);
      return false;
    }
  };

  // Auto-select used range
  const handleAutoSelectRange = async () => {
    try {
      setLoading(true);
      setStatusMessage("Auto-selecting used range...");

      // Call the selectUsedRange method from the excelService
      const result = await excelService.selectUsedRange();
      
      if (result.success) {
        // Update UI to show selection is active
        setUseSelection(true);
        
        // Update the current selection with the result
        setCurrentSelection(result.data);
        
        // Refresh the selection data
        await refreshSelectionData();
        
        // Show a success message
        setStatusMessage(`Selected ${result.data.rowCount} rows x ${result.data.columnCount} columns`);
        setTimeout(() => setStatusMessage(null), 3000);
      } else {
        // If there was an error, show it
        setStatusMessage(`Error: ${result.error || 'Failed to select used range'}`);
        setTimeout(() => setStatusMessage(null), 3000);
      }
    } catch (error) {
      console.error("Error auto-selecting used range:", error);
      setStatusMessage(`Error: ${error.message || 'Failed to select used range'}`);
      setTimeout(() => setStatusMessage(null), 3000);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className={classes.root}>
      <div className={classes.chatContainer}>
        <ChatHeader 
          mode={mode} 
          onModeChange={handleModeChange}
          onClearChat={clearMessages}
          canUndo={canUndo}
          onUndo={handleUndo}
        />
        
        {renderChatContent()}
        
        <ChatInput 
          value={inputValue}
          onChange={handleInputChange}
          onKeyDown={handleKeyDown}
          onSend={handleSendMessage}
          disabled={isProcessing || processingAction}
          currentSelection={useSelection ? currentSelection : null}
          onToggleSelection={handleToggleSelection}
          onAutoSelectRange={handleAutoSelectRange}
        />
        
        {/* Status message pop-up */}
        {statusMessage && (
          <div className={classes.statusMessage}>
            {statusMessage}
          </div>
        )}
      </div>
      
      {/* Data View Button (only shown when selection is active) */}
      <ViewDataButton 
        selectionData={selectionData} 
        formatDataForDisplay={formatDataForDisplay}
        hasSelection={useSelection && !!currentSelection}
        openAIJsonData={openAIJsonData}
      />
    </div>
  );
} 