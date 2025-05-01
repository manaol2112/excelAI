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
    dispatch({ type: 'ADD_MESSAGE', payload: message });
  }, []);
  
  /**
   * Update the last message in the conversation
   */
  const updateLastMessage = useCallback((updates) => {
    dispatch({ type: 'UPDATE_LAST_MESSAGE', payload: updates });
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
  const loadDataProfile = useCallback(async (range = null) => {
    try {
      setError(null);
      const result = await dataAnalysisService.generateDataProfile(range);
      
      if (result.success) {
        dispatch({ 
          type: 'SET_DATA_PROFILE', 
          payload: result.profile 
        });
        return result.profile;
      } else {
        setError(result.error || 'Failed to analyze data');
        return null;
      }
    } catch (err) {
      setError(err.message || 'Error analyzing data');
      return null;
    }
  }, []);

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
    addMessage,
    updateLastMessage,
    clearMessages,
    messages: conversationState.messages,
    excelContext: conversationState.excelContext,
    dataProfile: conversationState.dataProfile,
    // Model selection
    selectedModel,
    setSelectedModel,
    selectedProvider,
    setSelectedProvider,
    availableModels: AI_MODELS
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