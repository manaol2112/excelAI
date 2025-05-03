import React, { createContext, useContext, useReducer, useState, useEffect, useCallback } from 'react';
import excelMiddleware from '../services/excelMiddleware';
import dataAnalysisService from '../services/dataAnalysisService';
import openaiService from '../services/openaiService';
import excelService from '../services/excelService';

// Define available AI models
const AI_MODELS = {
  OPENAI: [
    { id: 'gpt-3.5-turbo', name: 'GPT-3.5 Turbo', provider: 'OpenAI' },
    { id: 'gpt-4', name: 'GPT-4', provider: 'OpenAI' },
    { id: 'gpt-4-turbo', name: 'GPT-4 Turbo', provider: 'OpenAI' }
  ]
};

// Define conversation reducer for better state management
const conversationReducer = (state, action) => {
  switch (action.type) {
    case 'ADD_MESSAGE':
      return {
        ...state,
        messages: [...state.messages, action.payload],
        lastActivity: Date.now()
      };
    case 'UPDATE_LAST_MESSAGE':
      const updatedMessages = [...state.messages];
      if (updatedMessages.length > 0) {
        updatedMessages[updatedMessages.length - 1] = {
          ...updatedMessages[updatedMessages.length - 1],
          ...action.payload
        };
      }
      return {
        ...state,
        messages: updatedMessages,
        lastActivity: Date.now()
      };
    case 'CLEAR_MESSAGES':
      return {
        ...state,
        messages: [],
        lastActivity: Date.now()
      };
    case 'SET_EXCEL_CONTEXT':
      return {
        ...state,
        excelContext: action.payload,
        lastContextUpdate: Date.now()
      };
    case 'SET_DATA_PROFILE':
      return {
        ...state,
        dataProfile: action.payload,
        lastDataProfileUpdate: Date.now()
      };
    default:
      return state;
  }
};

// Initial state for the conversation
const initialConversationState = {
  messages: [],
  lastActivity: null,
  excelContext: null,
  lastContextUpdate: null,
  dataProfile: null,
  lastDataProfileUpdate: null
};

// Create the context
const AIContext = createContext();

// Create the provider component
export const AIProvider = ({ children }) => {
  const [apiKey, setApiKey] = useState(() => {
    const savedKey = localStorage.getItem('openai_api_key');
    return savedKey || '';
  });
  
  // Add isApiKeyValid state
  const [isApiKeyValid, setIsApiKeyValid] = useState(!!apiKey);
  
  // Model selection state
  const [selectedProvider, setSelectedProvider] = useState('OPENAI');
  const [selectedModel, setSelectedModel] = useState(() => {
    const savedModel = localStorage.getItem('excelai_selected_model');
    if (savedModel) {
      try {
        return JSON.parse(savedModel);
      } catch (e) {
        console.error('Error parsing saved model:', e);
        return AI_MODELS.OPENAI[0];
      }
    }
    return AI_MODELS.OPENAI[0];
  });
  
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState(null);
  const [undoStack, setUndoStack] = useState([]);
  const [canUndo, setCanUndo] = useState(false);
  
  // Use reducer for conversation state management
  const [conversationState, dispatch] = useReducer(
    conversationReducer,
    initialConversationState
  );
  
  // Save API key to local storage and update the OpenAI service
  useEffect(() => {
    localStorage.setItem('openai_api_key', apiKey);
    
    // Update OpenAI service with new key and validate it
    if (apiKey) {
      try {
        openaiService.setApiKey(apiKey);
        setIsApiKeyValid(true);
      } catch (error) {
        console.error('Error setting API key:', error);
        setIsApiKeyValid(false);
      }
    } else {
      setIsApiKeyValid(false);
    }
  }, [apiKey]);

  // Save selected model when it changes
  useEffect(() => {
    localStorage.setItem('excelai_selected_model', JSON.stringify(selectedModel));
  }, [selectedModel]);

  // Check undo capability
  useEffect(() => {
    setCanUndo(undoStack.length > 0);
  }, [undoStack]);
  
  /**
   * Add a new message to the conversation
   */
  const addMessage = useCallback((message) => {
    try {
      // Ensure message content is a string if present
      const safeMessage = { ...message };
      if (safeMessage.content && typeof safeMessage.content !== 'string') {
        safeMessage.content = String(safeMessage.content);
      }
      dispatch({ type: 'ADD_MESSAGE', payload: safeMessage });
    } catch (error) {
      console.error("Error adding message:", error);
    }
  }, []);
  
  /**
   * Update the last message in the conversation
   */
  const updateLastMessage = useCallback((updates) => {
    try {
      // Ensure content is a string if present
      const safeUpdates = { ...updates };
      if (safeUpdates.content && typeof safeUpdates.content !== 'string') {
        safeUpdates.content = String(safeUpdates.content);
      }
      dispatch({ type: 'UPDATE_LAST_MESSAGE', payload: safeUpdates });
    } catch (error) {
      console.error("Error updating last message:", error);
    }
  }, []);
  
  /**
   * Clear all messages in the conversation
   */
  const clearMessages = useCallback(() => {
    dispatch({ type: 'CLEAR_MESSAGES' });
  }, []);
  
  /**
   * Load Excel context and update state
   */
  const loadExcelContext = useCallback(async (forceRefresh = false) => {
    try {
      setError(null);
      const result = await excelMiddleware.getWorkbookState(forceRefresh);
      
      if (result.success) {
        dispatch({ 
          type: 'SET_EXCEL_CONTEXT', 
          payload: result.data 
        });
        return result.data;
      } else {
        setError(result.error || 'Failed to load Excel context');
        return null;
      }
    } catch (err) {
      setError(err.message || 'Error loading Excel context');
      return null;
    }
  }, []);
  
  /**
   * Load data profile for the current selection or range
   */
  const loadDataProfile = async () => {
    try {
      setIsProcessing(true);
      // Load basic worksheet data first
      const workbookState = await excelService.preloadWorkbookData(true);
      
      if (!workbookState) {
        setIsProcessing(false);
        return null;
      }
      
      // Get more detailed analysis of data
      const usedRangeData = await excelService.getAllData();
      
      setIsProcessing(false);
      return {
        success: true,
        workbookState,
        usedRangeData
      };
    } catch (error) {
      setIsProcessing(false);
      setError(`Error loading data profile: ${error.message}`);
      return null;
    }
  };
  
  /**
   * Analyze data from a specific selection range
   * @param {string} selectionAddress - The address of the selection to analyze
   * @returns {Promise<Object>} Analysis results for the selection
   */
  const getSelectionData = async (selectionAddress) => {
    try {
      setIsProcessing(true);
      console.log(`Starting analysis of range: ${selectionAddress}`);
      
      // Check if we have a valid selection address
      if (!selectionAddress) {
        console.error('No selection address provided');
        setIsProcessing(false);
        return { success: false, error: 'No selection address provided' };
      }
      
      // Validate selection address format (simple regex check)
      if (!/^[a-zA-Z]+\d+:[a-zA-Z]+\d+$|^[a-zA-Z]+\d+$/.test(selectionAddress)) {
        console.error(`Invalid selection address format: ${selectionAddress}`);
        setIsProcessing(false);
        return { success: false, error: `Invalid selection address format: ${selectionAddress}` };
      }
      
      // First refresh Excel context to ensure we have up-to-date data
      const contextResult = await loadExcelContext(true);
      if (!contextResult) {
        console.error('Failed to load Excel context before selection analysis');
        setIsProcessing(false);
        return { success: false, error: 'Failed to load Excel context' };
      }
      
      console.log(`Excel context loaded, proceeding to analyze selection: ${selectionAddress}`);
      console.log(`Context contains selection: ${contextResult.selection ? 'Yes' : 'No'}`);
      if (contextResult.selection) {
        console.log(`Context selection address: ${contextResult.selection.address}`);
      }
      
      // Now get the selection data
      const result = await excelService.getSelectionData(selectionAddress);
      
      // Validate the result
      if (result && result.success && result.data) {
        // Additional validation on the returned data
        if (!result.data.values || !Array.isArray(result.data.values)) {
          console.error('Selection data does not contain valid values array');
          setIsProcessing(false);
          return { success: false, error: 'Selection data invalid: No values array found' };
        }
        
        // Log data sizes for debugging
        console.log(`Successfully analyzed selection: ${selectionAddress} with dimensions:
          - Rows: ${result.data.summary.rowCount}
          - Columns: ${result.data.summary.columnCount}
          - Non-empty cells: ${result.data.summary.nonEmptyCells}
          - Empty cells: ${result.data.summary.emptyCells}
          - Contains numeric data: ${result.data.summary.hasNumericData ? 'Yes' : 'No'}
          - Contains text data: ${result.data.summary.hasTextData ? 'Yes' : 'No'}
        `);
        
        // Sample value from the result for verification (just first few cells)
        if (result.data.values.length > 0) {
          const sampleValues = result.data.values.slice(0, Math.min(2, result.data.values.length))
            .map(row => row.slice(0, Math.min(3, row.length)));
          console.log('Sample values from selection:', JSON.stringify(sampleValues));
        }
      } else {
        console.error('Failed to get selection data:', result?.error || 'Unknown error');
      }
      
      setIsProcessing(false);
      return result;
    } catch (error) {
      console.error(`Error analyzing selection: ${error.message}`, error);
      setIsProcessing(false);
      setError(`Error analyzing selection: ${error.message}`);
      return { success: false, error: error.message };
    }
  };

  /**
   * Build a comprehensive context for AI prompts
   */
  const buildAIContext = useCallback(async (requestType = 'general', forceRefresh = false) => {
    try {
      // Load Excel context if needed
      const needsContextRefresh = forceRefresh || 
        !conversationState.excelContext || 
        (Date.now() - conversationState.lastContextUpdate > 30000); // 30 seconds
      
      let excelContext = conversationState.excelContext;
      if (needsContextRefresh) {
        excelContext = await loadExcelContext(forceRefresh);
      }
      
      // For data analysis, load data profile
      let dataProfile = null;
      if (requestType === 'analyze' || requestType === 'chart') {
        // Check if we need to refresh the data profile
        const needsProfileRefresh = forceRefresh || 
          !conversationState.dataProfile || 
          (Date.now() - conversationState.lastDataProfileUpdate > 30000); // 30 seconds
        
        if (needsProfileRefresh) {
          dataProfile = await loadDataProfile();
        } else {
          dataProfile = conversationState.dataProfile;
        }
      }
      
      return {
        excelContext,
        dataProfile,
        messages: conversationState.messages,
        requestType
      };
    } catch (err) {
      setError(err.message || 'Error building AI context');
      return null;
    }
  }, [loadExcelContext, loadDataProfile, conversationState]);
  
  /**
   * Generate text using OpenAI with Excel context
   */
  const generateText = useCallback(async (prompt, requestType = 'general', options = {}) => {
    try {
      setIsProcessing(true);
      setError(null);
      
      // Build context for AI
      const context = await buildAIContext(requestType, options.forceRefresh);
      
      if (!context) {
        throw new Error('Failed to build context for AI');
      }
      
      // Prepare conversation history from messages
      const conversationHistory = conversationState.messages.map(msg => ({
        role: msg.role,
        content: msg.content
      }));
      
      // Create prompt data with context
      const promptData = {
        prompt,
        excelContext: context.excelContext,
        dataProfile: context.dataProfile,
        requestType,
        conversationHistory,
        systemDirectives: options.systemDirectives || null
      };
      
      // Call OpenAI service with selected model ID
      const result = await openaiService.generateText(promptData, selectedModel.id);
      
      setIsProcessing(false);
      return result;
    } catch (err) {
      setIsProcessing(false);
      setError(err.message || 'Error generating AI response');
      throw err;
    }
  }, [buildAIContext, conversationState.messages, selectedModel]);
  
  /**
   * Generate a chart suggestion
   */
  const suggestChart = useCallback(async (prompt, options = {}) => {
    try {
      setIsProcessing(true);
      setError(null);
      
      // Build context for AI
      const context = await buildAIContext('chart', options.forceRefresh);
      
      if (!context) {
        throw new Error('Failed to build context for chart suggestion');
      }
      
      // Call OpenAI service for chart suggestion with selected model
      const result = await openaiService.suggestChart(prompt, context, selectedModel.id);
      
      setIsProcessing(false);
      return result;
    } catch (err) {
      setIsProcessing(false);
      setError(err.message || 'Error generating chart suggestion');
      throw err;
    }
  }, [buildAIContext, selectedModel]);
  
  /**
   * Generate a formula suggestion
   */
  const suggestFormula = useCallback(async (prompt, options = {}) => {
    try {
      setIsProcessing(true);
      setError(null);
      
      // Build context for AI
      const context = await buildAIContext('formula', options.forceRefresh);
      
      if (!context) {
        throw new Error('Failed to build context for formula suggestion');
      }
      
      // Call OpenAI service for formula suggestion with selected model
      const result = await openaiService.suggestFormula(prompt, context, selectedModel.id);
      
      setIsProcessing(false);
      return result;
    } catch (err) {
      setIsProcessing(false);
      setError(err.message || 'Error generating formula suggestion');
      throw err;
    }
  }, [buildAIContext, selectedModel]);
  
  /**
   * Perform data analysis
   */
  const analyzeData = useCallback(async (prompt, options = {}) => {
    try {
      setIsProcessing(true);
    setError(null);
      
      // Build context for AI
      const context = await buildAIContext('analyze', options.forceRefresh);
      
      if (!context) {
        throw new Error('Failed to build context for data analysis');
      }
      
      // Call OpenAI service for data analysis with selected model
      const result = await openaiService.analyzeData(prompt, context, selectedModel.id);
      
      setIsProcessing(false);
      return result;
    } catch (err) {
      setIsProcessing(false);
      setError(err.message || 'Error analyzing data');
      throw err;
    }
  }, [buildAIContext, selectedModel]);
  
  /**
   * Execute Excel code (with undo tracking)
   */
  const executeExcelCode = useCallback(async (code, trackUndo = true) => {
    try {
      setIsProcessing(true);
      setError(null);
      
      // Track state for undo if needed
      let preState = null;
      if (trackUndo) {
        preState = await excelService.getCurrentState();
      }
      
      // Execute the code
      const result = await excelService.execute(code);
      
      // Add to undo stack if successful and tracking
      if (result.success && trackUndo && preState) {
        setUndoStack(prev => [...prev, preState]);
      }
      
      setIsProcessing(false);
      return result;
    } catch (err) {
      setIsProcessing(false);
      setError(err.message || 'Error executing Excel code');
      throw err;
    }
  }, []);
  
  /**
   * Count occurrences of a value in a column
   * @param {string} column - Column letter to analyze
   * @param {string} value - Value to count
   * @param {boolean} caseInsensitive - Whether to ignore case
   * @returns {Promise<Object>} - Count results
   */
  const countValueInColumn = useCallback(async (column, value, caseInsensitive = true) => {
    try {
      setIsProcessing(true);
      setError(null);
      
      const result = await excelService.countInColumn(column, value, caseInsensitive);
      
      setIsProcessing(false);
      return result;
    } catch (err) {
      setIsProcessing(false);
      setError(err.message || `Error counting "${value}" in column ${column}`);
      throw err;
    }
  }, []);

  /**
   * Get detailed analysis of a column
   * @param {string} column - Column letter to analyze
   * @returns {Promise<Object>} - Analysis results
   */
  const analyzeColumnData = useCallback(async (column) => {
    try {
      setIsProcessing(true);
      setError(null);
      
      const result = await excelService.analyzeColumn(column);
      
      setIsProcessing(false);
      return result;
    } catch (err) {
      setIsProcessing(false);
      setError(err.message || `Error analyzing column ${column}`);
      throw err;
    }
  }, []);
  
  /**
   * Get analysis data preview
   * @param {string} range - The range to analyze
   * @param {number} maxRows - The maximum number of rows to include in the preview
   * @returns {Promise<Object>} - Analysis data preview
   */
  const getAnalysisDataPreview = useCallback(async (range = null, maxRows = 10) => {
    const profileResult = await dataAnalysisService.generateDataProfile(range);
    if (!profileResult.success) return { success: false, error: profileResult.error };
    const { profile } = profileResult;
    // Get headers
    const headers = profile.columns.map(col => col.name);
    // Get data rows (excluding header if present)
    const rangeContext = await excelMiddleware.extractDataAnalysisContext(range);
    if (!rangeContext.success) return { success: false, error: rangeContext.error };
    let dataRows = rangeContext.data.data;
    if (profile.hasHeaders) dataRows = dataRows.slice(1);
    // Build array of objects for preview
    const preview = dataRows.slice(0, maxRows).map(row =>
      Object.fromEntries(headers.map((h, i) => [h, row[i]]))
    );
    return { success: true, headers, preview, totalRows: dataRows.length };
  }, []);
  
  /**
   * Directly answer data questions by executing analysis on Excel data
   * @param {string} question - The user's question
   * @returns {Promise<Object>} - The answer
   */
  const answerDataQuestion = useCallback(async (question) => {
    try {
      setIsProcessing(true);
      setError(null);

      // 0. Data preview: "show me the data used", "what data was analyzed", etc.
      if (/show (me )?(the )?data( used)?( in (this|the) analysis)?/i.test(question) || /what data was analyzed/i.test(question)) {
        const previewResult = await getAnalysisDataPreview(null, 20);
        setIsProcessing(false);
        if (previewResult.success) {
          return {
            success: true,
            answer: `Here is a preview of the data used in the analysis (showing up to 20 rows):`,
            headers: previewResult.headers,
            preview: previewResult.preview,
            totalRows: previewResult.totalRows
          };
        } else {
          return { success: false, error: previewResult.error };
        }
      }

      // 1. Multi-criteria count (already handled above)
      const multiCriteriaPattern = /how many\s+([\w\s]+?)(?:\s+status)?(?:\s+with|\s+where)?\s+(.+)/i;
      const match = question.match(multiCriteriaPattern);
      if (match && match[1] && match[2]) {
        // Main value and column (e.g., 'Open status')
        let mainValue = match[1].trim();
        let mainColumn = null;
        // Try to infer main column
        const possibleColumns = ['status', 'state', 'condition', 'stage'];
        for (const col of possibleColumns) {
          if (question.toLowerCase().includes(col)) {
            mainColumn = col.charAt(0).toUpperCase() + col.slice(1);
            break;
          }
        }
        // Parse additional criteria (split by 'and')
        const criteria = [];
        if (mainColumn) {
          criteria.push({ column: mainColumn, value: mainValue, op: '=' });
        }
        const rest = match[2];
        const conds = rest.split(/\s+and\s+/i);
        for (const cond of conds) {
          // Match patterns like 'Amount > 1000', 'Region is Europe', 'Score >= 90'
          const condMatch = cond.match(/([\w\s]+?)\s*(=|is|>|<|>=|<=)\s*([\w\s\.\-]+)/i);
          if (condMatch) {
            let col = condMatch[1].trim();
            let op = condMatch[2].replace('is', '=').replace('=', '=');
            let val = condMatch[3].trim();
            // Try to parse number if possible
            if (!isNaN(val) && val !== '') val = Number(val);
            criteria.push({ column: col, value: val, op });
          }
        }
        if (criteria.length > 0) {
          const result = await dataAnalysisService.countRowsWithCriteria(criteria);
          setIsProcessing(false);
          if (result && typeof result.count === 'number') {
            return {
              success: true,
              answer: `There ${result.count === 1 ? 'is' : 'are'} ${result.count} row${result.count === 1 ? '' : 's'} matching all criteria.`,
              count: result.count,
              rowIndices: result.rowIndices
            };
          }
        }
      }

      // 2. Aggregation questions (sum, avg, min, max)
      // Patterns: 'total unit sold by Mark', 'sum of sales for Alice', 'average sales for Bob', 'min score for Alice', etc.
      const aggPatterns = [
        { type: 'sum', regex: /(total|sum(?: of)?|add(?: up)?)([\w\s]+?)(?:by|for|of)?\s+([\w\s]+)?\??$/i },
        { type: 'avg', regex: /(average|mean)([\w\s]+?)(?:by|for|of)?\s+([\w\s]+)?\??$/i },
        { type: 'min', regex: /(minimum|min|lowest)([\w\s]+?)(?:by|for|of)?\s+([\w\s]+)?\??$/i },
        { type: 'max', regex: /(maximum|max|highest)([\w\s]+?)(?:by|for|of)?\s+([\w\s]+)?\??$/i }
      ];
      for (const agg of aggPatterns) {
        const aggMatch = question.match(agg.regex);
        if (aggMatch) {
          // Try to extract target column and filter value
          let targetColumn = aggMatch[2] ? aggMatch[2].replace(/^(of|for|by)\s+/i, '').trim() : null;
          let filterValue = aggMatch[3] ? aggMatch[3].trim() : null;
          let filterColumn = null;
          // Try to infer filter column (e.g., Name, Sales Rep, etc.)
          const possibleFilterColumns = ['name', 'sales rep', 'salesperson', 'employee', 'person', 'rep'];
          for (const col of possibleFilterColumns) {
            if (question.toLowerCase().includes(col)) {
              filterColumn = col.split(' ').map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(' ');
              break;
            }
          }
          // If not found, try to guess from context (could be improved with LLM)
          if (!filterColumn && filterValue) {
            filterColumn = 'Name'; // Default guess
          }
          // Clean up target column
          if (targetColumn) {
            targetColumn = targetColumn.replace(/\bby\b|\bfor\b|\bof\b/gi, '').trim();
          }
          // Build criteria
          const criteria = [];
          if (filterColumn && filterValue) {
            criteria.push({ column: filterColumn, value: filterValue, op: '=' });
          }
          if (targetColumn) {
            const result = await dataAnalysisService.aggregateColumnWithCriteria(agg.type, targetColumn, criteria);
            setIsProcessing(false);
            if (result && typeof result.result === 'number') {
              let aggWord = agg.type === 'sum' ? 'total' : agg.type === 'avg' ? 'average' : agg.type === 'min' ? 'minimum' : 'maximum';
              return {
                success: true,
                answer: `The ${aggWord} ${targetColumn}${criteria.length > 0 ? ` for ${filterValue}` : ''} is ${result.result}.`,
                value: result.result,
                count: result.count,
                rowIndices: result.rowIndices
              };
            }
          }
        }
      }

      // Fallback to previous logic (single-criteria count, etc.)
      // Try to extract column from question
      const columnMatch = question.match(/column\s+([a-z])/i);
      let column = null;
      if (columnMatch && columnMatch[1]) {
        column = columnMatch[1].toUpperCase();
      }
      
      // Try to extract target values like "Closed", "Open", etc.
      const statusMatch = question.match(/(open|closed|pending|completed|in progress)/i);
      let targetValue = null;
      if (statusMatch && statusMatch[1]) {
        targetValue = statusMatch[1].charAt(0).toUpperCase() + statusMatch[1].slice(1).toLowerCase();
      }
      
      let result = null;
      
      // If we have both column and value, do a count
      if (column && targetValue) {
        const countResult = await excelService.countInColumn(column, targetValue, true);
        if (countResult.success) {
          const answer = `There ${countResult.count === 1 ? 'is' : 'are'} ${countResult.count} ${targetValue} status item${countResult.count === 1 ? '' : 's'} in column ${column}.`;
          result = {
            success: true,
            answer,
            data: countResult
          };
        }
      } 
      // If we just have a column, analyze it
      else if (column) {
        const analysisResult = await excelService.analyzeColumn(column);
        if (analysisResult.success && analysisResult.analysis) {
          const analysis = analysisResult.analysis;
          
          let answer = `Column ${column} contains ${analysis.nonEmptyCells} non-empty cells out of ${analysis.totalCells} total cells.`;
          
          // Add information about top values
          if (analysis.topValues && analysis.topValues.length > 0) {
            answer += ` The most common value is "${analysis.topValues[0].value}" which appears ${analysis.topValues[0].count} time${analysis.topValues[0].count === 1 ? '' : 's'}.`;
          }
          
          // Add numerical stats if available
          if (analysis.numericalStats) {
            const stats = analysis.numericalStats;
            answer += ` The column contains numerical data with an average of ${stats.average.toFixed(2)}, ranging from ${stats.min} to ${stats.max}.`;
          }
          
          result = {
            success: true,
            answer,
            data: analysisResult
          };
        }
      }
      
      // If we couldn't answer directly, fall back to AI analysis
      if (!result) {
        const aiContext = await buildAIContext('analyze', true);
        result = await openaiService.analyzeData(question, aiContext, selectedModel.id);
      }
      
      setIsProcessing(false);
      return result;
    } catch (err) {
      setIsProcessing(false);
      setError(err.message || 'Error answering data question');
      throw err;
    }
  }, [buildAIContext, selectedModel]);
  
  /**
   * Undo the last Excel operation
   */
  const undoLastOperation = useCallback(async () => {
    if (undoStack.length === 0) {
      return { success: false, error: 'Nothing to undo' };
    }
    
    try {
      setIsProcessing(true);
      setError(null);
      
      // Get the last state from the stack
      const lastState = undoStack[undoStack.length - 1];
      
      // Apply the state
      const result = await excelService.restoreState(lastState);
      
      // Remove from undo stack
      setUndoStack(prev => prev.slice(0, -1));
      
      setIsProcessing(false);
      return result;
    } catch (err) {
      setIsProcessing(false);
      setError(err.message || 'Error undoing last operation');
      throw err;
    }
  }, [undoStack]);

  /**
   * Count occurrences of a specific value in a specified range
   * @param {string} rangeAddress - The range address to search in
   * @param {string} targetValue - The value to count
   * @param {boolean} caseSensitive - Whether to match case-sensitively
   * @returns {Promise<Object>} Count result with location information
   */
  const countValueInRange = useCallback(async (rangeAddress, targetValue, caseSensitive = false) => {
    try {
      setIsProcessing(true);
      setError(null);
      console.log(`AIContext: Counting "${targetValue}" in range: ${rangeAddress}`);
      
      const result = await excelService.countValueInRange(rangeAddress, targetValue, caseSensitive);
      
      setIsProcessing(false);
      return result;
    } catch (err) {
      setIsProcessing(false);
      setError(err.message || `Error counting "${targetValue}" in range ${rangeAddress}`);
      return { success: false, error: err.message };
    }
  }, []);

  // Context value
  const value = {
    apiKey,
    setApiKey,
    isApiKeyValid,
    isProcessing,
    error,
    canUndo,
    generateText,
    suggestFormula,
    suggestChart,
    analyzeData,
    executeExcelCode,
    undoLastOperation,
    loadExcelContext,
    loadDataProfile,
    getSelectionData,
    addMessage,
    updateLastMessage,
    clearMessages,
    messages: conversationState.messages,
    excelContext: conversationState.excelContext,
    dataProfile: conversationState.dataProfile,
    // Direct data analysis functions
    countValueInColumn,
    countValueInRange,
    analyzeColumnData,
    answerDataQuestion,
    // Model selection
    selectedModel,
    setSelectedModel,
    selectedProvider,
    setSelectedProvider,
    availableModels: AI_MODELS,
    excelService
  };

  return <AIContext.Provider value={value}>{children}</AIContext.Provider>;
};

// Custom hook for using the AI context
export const useAI = () => {
  const context = useContext(AIContext);
  
  if (!context) {
    throw new Error('useAI must be used within an AIProvider');
  }
  
  return context;
}; 

export default AIContext; 