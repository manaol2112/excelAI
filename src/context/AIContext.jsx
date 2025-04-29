import React, { createContext, useState, useContext, useEffect } from 'react';
import OpenAIService from '../services/openaiService';
import excelService from '../services/excelService';

// Default models available
const AI_MODELS = {
  OPENAI: [
    { id: 'gpt-3.5-turbo', name: 'GPT-3.5 Turbo', provider: 'OpenAI' },
    { id: 'gpt-4', name: 'GPT-4', provider: 'OpenAI' },
    { id: 'gpt-4-turbo', name: 'GPT-4 Turbo', provider: 'OpenAI' }
  ]
};

// Create context
const AIContext = createContext();

export const AIProvider = ({ children }) => {
  // State for API key
  const [apiKey, setApiKey] = useState('');
  const [isApiKeyValid, setIsApiKeyValid] = useState(false);

  // State for selected model
  const [selectedProvider, setSelectedProvider] = useState('OPENAI');
  const [selectedModel, setSelectedModel] = useState(AI_MODELS.OPENAI[0]);
  
  // OpenAI service instance
  const [aiService, setAiService] = useState(null);
  
  // Loading state
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState(null);

  // State for tracking the execution environment
  const [macrosEnabled, setMacrosEnabled] = useState(false);
  const [executionHistory, setExecutionHistory] = useState([]);
  
  // Token usage tracking
  const [tokenUsageStats, setTokenUsageStats] = useState(null);

  // Initialize or update AI service when API key changes
  useEffect(() => {
    if (apiKey) {
      try {
        const service = new OpenAIService(apiKey);
        setAiService(service);
        // We'll consider the key valid by default when set
        // In a production app, you'd validate this key more thoroughly
        setIsApiKeyValid(true);
        setError(null);
        
        // Store API key in localStorage
        localStorage.setItem('excelai_api_key', apiKey);
      } catch (err) {
        setError('Failed to initialize AI service');
        setIsApiKeyValid(false);
        console.error('Error initializing AI service:', err);
      }
    } else {
      setAiService(null);
      setIsApiKeyValid(false);
    }
  }, [apiKey]);

  // Load saved API key on initial render
  useEffect(() => {
    const savedApiKey = localStorage.getItem('excelai_api_key');
    if (savedApiKey) {
      setApiKey(savedApiKey);
    }
    
    // Load saved model preference
    const savedModel = localStorage.getItem('excelai_selected_model');
    if (savedModel) {
      try {
        const modelObj = JSON.parse(savedModel);
        setSelectedModel(modelObj);
        setSelectedProvider(modelObj.provider === 'OpenAI' ? 'OPENAI' : modelObj.provider);
      } catch (e) {
        console.error('Error parsing saved model:', e);
      }
    }

    // Load macro setting preference
    const macroSetting = localStorage.getItem('excelai_macros_enabled');
    if (macroSetting) {
      setMacrosEnabled(macroSetting === 'true');
    }

    // Detect if we're in an environment that supports macros
    checkMacroSupport();
  }, []);

  // Save selected model when it changes
  useEffect(() => {
    localStorage.setItem('excelai_selected_model', JSON.stringify(selectedModel));
  }, [selectedModel]);

  // Save macro enabled setting when it changes
  useEffect(() => {
    localStorage.setItem('excelai_macros_enabled', macrosEnabled.toString());
  }, [macrosEnabled]);

  // Update token usage stats periodically
  useEffect(() => {
    if (aiService) {
      const updateStats = () => {
        const stats = aiService.getTokenUsageStats();
        setTokenUsageStats(stats);
      };
      
      // Update initially
      updateStats();
      
      // Update every 30 seconds
      const interval = setInterval(updateStats, 30000);
      
      return () => clearInterval(interval);
    }
  }, [aiService]);

  // Check if the current environment supports macros
  const checkMacroSupport = async () => {
    // This is a basic check - in a real app, you might have a more robust detection
    try {
      if (typeof Office !== 'undefined' && Office.context) {
        // Check if we're in desktop Excel (which can support macros)
        if (Office.context.host === 'Excel' && 
            (Office.context.platform === 'PC' || Office.context.platform === 'Mac')) {
          setMacrosEnabled(true);
        }
      }
    } catch (err) {
      console.error('Error checking macro support:', err);
      setMacrosEnabled(false);
    }
  };

  // Function to handle API requests with loading state
  const callAIWithLoading = async (apiFunction, ...args) => {
    setIsLoading(true);
    setError(null);
    try {
      if (!aiService) {
        throw new Error('AI service not initialized. Please enter your API key.');
      }
      const result = await apiFunction.apply(aiService, [...args, selectedModel.id]);
      
      // Update token usage stats after each call
      const stats = aiService.getTokenUsageStats();
      setTokenUsageStats(stats);
      
      return result;
    } catch (err) {
      setError(err.message || 'An error occurred while calling the AI service');
      console.error('AI service error:', err);
      throw err;
    } finally {
      setIsLoading(false);
    }
  };

  // Functions to call AI methods with loading state
  const generateText = async (prompt) => {
    return callAIWithLoading(aiService?.generateText, prompt);
  };

  const suggestFormula = async (description) => {
    return callAIWithLoading(aiService?.suggestFormula, description);
  };

  const analyzeData = async (data, analysisType) => {
    return callAIWithLoading(aiService?.analyzeData, data, analysisType);
  };

  const generateChart = async (data, chartType) => {
    return callAIWithLoading(aiService?.generateChart, data, chartType);
  };

  // New function to determine if a task requires VBA macros
  const checkMacroRequirement = async (task) => {
    return callAIWithLoading(aiService?.detectMacroRequirement, task);
  };

  // New function to generate VBA code for tasks that require macros
  const generateVBACode = async (task) => {
    return callAIWithLoading(aiService?.generateVBACode, task);
  };

  // Function to reset token usage statistics
  const resetTokenUsageStats = () => {
    if (aiService) {
      aiService.resetTokenUsageStats();
      setTokenUsageStats(aiService.getTokenUsageStats());
    }
  };

  // Calculate estimated cost based on token usage and selected model
  const calculateEstimatedCost = () => {
    if (!tokenUsageStats) return { usd: 0 };
    
    // Define cost per 1000 tokens for different models (input/output)
    const modelCosts = {
      'gpt-3.5-turbo': { input: 0.0015, output: 0.002 },
      'gpt-4': { input: 0.03, output: 0.06 },
      'gpt-4-turbo': { input: 0.01, output: 0.03 }
    };
    
    // Get cost for current model
    const modelId = selectedModel.id;
    const cost = modelCosts[modelId] || modelCosts['gpt-3.5-turbo']; // default to 3.5-turbo costs
    
    // Estimate total tokens (incomplete since we only track completion tokens explicitly)
    // For a rough estimate, we assume input tokens are around 1.2x output tokens
    const outputTokens = tokenUsageStats.totalTokensUsed;
    const estimatedInputTokens = Math.round(outputTokens * 1.2);
    
    // Calculate cost
    const inputCost = (estimatedInputTokens / 1000) * cost.input;
    const outputCost = (outputTokens / 1000) * cost.output;
    const totalCost = inputCost + outputCost;
    
    return {
      usd: totalCost.toFixed(4),
      breakdown: {
        input: {
          tokens: estimatedInputTokens,
          cost: inputCost.toFixed(4)
        },
        output: {
          tokens: outputTokens,
          cost: outputCost.toFixed(4)
        }
      }
    };
  };

  // Function to execute AI-generated Office.js code
  const executeOfficeJsCode = async (code) => {
    setIsLoading(true);
    setError(null);
    try {
      // Strip the js code block markers if present
      let codeToExecute = code;
      if (codeToExecute.startsWith('```js')) {
        codeToExecute = codeToExecute.replace(/^```js\n/, '').replace(/```$/, '');
      } else if (codeToExecute.startsWith('```javascript')) {
        codeToExecute = codeToExecute.replace(/^```javascript\n/, '').replace(/```$/, '');
      }

      // Extract the operation type from the code for better feedback
      const operationType = extractOperationType(codeToExecute);

      // Execute the code
      const result = await excelService.executeOfficeJsCode(codeToExecute);
      
      // Track execution history
      setExecutionHistory(prev => [...prev, {
        timestamp: new Date(),
        type: 'office.js',
        code: codeToExecute,
        success: result.success,
        error: result.error
      }]);
      
      // Generate a human-friendly response message
      if (result.success) {
        result.humanMessage = generateSuccessMessage(operationType);
      } else {
        result.humanMessage = `The operation failed: ${result.error || 'Unknown error'}`;
      }
      
      return result;
    } catch (err) {
      setError(`Error executing Office.js code: ${err.message}`);
      console.error('Error executing Office.js code:', err);
      
      // Track failed execution
      setExecutionHistory(prev => [...prev, {
        timestamp: new Date(),
        type: 'office.js',
        code,
        success: false,
        error: err.message
      }]);
      
      // Add human-friendly error message
      const errorResult = { 
        success: false, 
        error: err.message,
        humanMessage: `Unable to complete the operation: ${err.message}`
      };
      
      throw errorResult;
    } finally {
      setIsLoading(false);
    }
  };

  // Helper function to extract the operation type from code
  const extractOperationType = (code) => {
    // Default operation type if we can't determine specifics
    let operationType = 'Excel operation';
    
    // Try to determine what kind of operation was performed
    if (/format\.fill\.color|background|\.fill\s*=/.test(code)) {
      operationType = 'cell coloring';
    } else if (/format\.font\.(bold|italic|underline|color)/.test(code)) {
      operationType = 'text formatting';
    } else if (/\.merge\(\)/.test(code)) {
      operationType = 'cell merging';
    } else if (/\.unmerge\(\)/.test(code)) {
      operationType = 'cell unmerging';
    } else if (/format\.(row|column)(Height|Width)\s*=\s*0/.test(code)) {
      operationType = 'hiding rows/columns';
    } else if (/format\.(row|column)(Height|Width)\s*=\s*[1-9]/.test(code)) {
      operationType = 'unhiding rows/columns';
    } else if (/charts\.add/.test(code)) {
      operationType = 'chart creation';
    } else if (/tables\.add/.test(code)) {
      operationType = 'table creation';
    } else if (/pivotTables\.add/.test(code)) {
      operationType = 'pivot table creation';
    } else if (/\.values\s*=/.test(code)) {
      operationType = 'data entry';
    } else if (/\.formulas\s*=/.test(code)) {
      operationType = 'formula insertion';
    } else if (/conditionalFormats\.add/.test(code)) {
      operationType = 'conditional formatting';
    } else if (/autoFilter\.apply/.test(code)) {
      operationType = 'filtering';
    } else if (/\.sort\.apply/.test(code)) {
      operationType = 'sorting';
    } else if (/worksheets\.add/.test(code)) {
      operationType = 'worksheet creation';
    } else if (/sheet\.name\s*=/.test(code)) {
      operationType = 'worksheet renaming';
    } else if (/protection\.protect/.test(code)) {
      operationType = 'worksheet protection';
    } else if (/protection\.unprotect/.test(code)) {
      operationType = 'worksheet unprotection';
    } else if (/dataValidation\.rule/.test(code)) {
      operationType = 'data validation';
    } else if (/hyperlink\s*=/.test(code)) {
      operationType = 'hyperlink insertion';
    } else if (/comments\.add/.test(code)) {
      operationType = 'comment addition';
    } else if (/range.select\(\)/.test(code)) {
      operationType = 'range selection';
    } else if (/slicers\.add/.test(code)) {
      operationType = 'slicer creation';
    } else if (/format\.autofitColumns/.test(code)) {
      operationType = 'column autofit';
    } else if (/format\.autofitRows/.test(code)) {
      operationType = 'row autofit';
    }
    
    return operationType;
  };

  // Helper function to generate success messages
  const generateSuccessMessage = (operationType) => {
    const successPhrases = [
      'Successfully completed',
      'Successfully performed',
      'Successfully executed',
      'Completed',
      'Finished',
      'Applied',
      'Done'
    ];
    
    // Pick a random success phrase
    const phrase = successPhrases[Math.floor(Math.random() * successPhrases.length)];
    
    return `${phrase} the ${operationType}.`;
  };

  // Function to execute a VBA macro
  const executeMacro = async (macroName, parameters = []) => {
    setIsLoading(true);
    setError(null);
    try {
      if (!macrosEnabled) {
        throw new Error('VBA macros are not enabled in this environment');
      }
      
      const result = await excelService.executeMacro(macroName, parameters);
      
      // Track execution history
      setExecutionHistory(prev => [...prev, {
        timestamp: new Date(),
        type: 'vba',
        macroName,
        parameters,
        success: result.success,
        error: result.error
      }]);
      
      // Add human-friendly message
      if (result.success) {
        result.humanMessage = `Successfully executed the macro "${macroName}".`;
      } else {
        result.humanMessage = `Failed to execute the macro "${macroName}": ${result.error || 'Unknown error'}`;
      }
      
      return result;
    } catch (err) {
      setError(`Error executing VBA macro: ${err.message}`);
      console.error('Error executing VBA macro:', err);
      
      // Track failed execution
      setExecutionHistory(prev => [...prev, {
        timestamp: new Date(),
        type: 'vba',
        macroName,
        parameters,
        success: false,
        error: err.message
      }]);
      
      // Add human-friendly error message
      const errorResult = { 
        success: false, 
        error: err.message,
        humanMessage: `Unable to execute the macro: ${err.message}`
      };
      
      throw errorResult;
    } finally {
      setIsLoading(false);
    }
  };

  // Function to process a user request and determine the execution path
  const processRequest = async (userRequest) => {
    setIsLoading(true);
    setError(null);
    try {
      // First, check if the task requires VBA macros
      const { requiresMacro, explanation } = await checkMacroRequirement(userRequest);
      
      if (requiresMacro) {
        // If macros are not enabled, inform the user
        if (!macrosEnabled) {
          return {
            success: false,
            requiresMacro: true,
            macrosEnabled: false,
            message: "This task requires VBA macros, but macros are not enabled in your current environment.",
            humanMessage: "This task requires VBA macros, but macros are not enabled in your current environment."
          };
        }
        
        // Generate VBA code for the task
        const vbaCode = await generateVBACode(userRequest);
        
        return {
          success: true,
          requiresMacro: true,
          macrosEnabled: true,
          vbaCode,
          explanation,
          humanMessage: "VBA code has been generated for your task. Would you like to execute it?"
        };
      } else {
        // The task can be done with Office.js - generate the code
        const officeJsCode = await generateText(userRequest);
        
        return {
          success: true,
          requiresMacro: false,
          officeJsCode,
          explanation,
          humanMessage: "Code has been generated for your task. Would you like to execute it?"
        };
      }
    } catch (err) {
      setError(`Error processing request: ${err.message}`);
      console.error('Error processing request:', err);
      return {
        success: false,
        error: err.message,
        humanMessage: `I couldn't process your request: ${err.message}`
      };
    } finally {
      setIsLoading(false);
    }
  };

  // Function to handle combined execution after processing
  const executeRequest = async (processedRequest) => {
    if (!processedRequest || !processedRequest.success) {
      return { 
        success: false, 
        error: processedRequest?.error || 'Invalid request processing result',
        humanMessage: processedRequest?.humanMessage || 'Sorry, I was unable to process your request.'
      };
    }
    
    let result;
    
    if (processedRequest.requiresMacro) {
      if (!processedRequest.macrosEnabled) {
        return { 
          success: false, 
          error: 'This task requires VBA macros, but macros are not enabled in your current environment.',
          humanMessage: 'This task requires VBA macros, but macros are not enabled in your current environment.'
        };
      }
      
      // Extract macro name and parameters from the VBA code
      // This is a placeholder - in reality, you would need a more sophisticated parser
      // or have the AI return the macro name and parameters separately
      const macroName = "ExecuteGeneratedMacro"; // Default name
      const parameters = [];
      
      result = await executeMacro(macroName, parameters);
    } else {
      // Execute Office.js code
      result = await executeOfficeJsCode(processedRequest.officeJsCode);
    }
    
    // Make sure there's always a human-friendly message
    if (!result.humanMessage) {
      if (result.success) {
        result.humanMessage = "The operation was completed successfully.";
      } else {
        result.humanMessage = `The operation failed: ${result.error || 'Unknown error'}`;
      }
    }
    
    return result;
  };

  // Handle model selection
  const handleModelSelect = (model) => {
    setSelectedModel(model);
  };

  // Toggle macro enabled setting
  const toggleMacrosEnabled = () => {
    setMacrosEnabled(prev => !prev);
  };

  // Context value
  const value = {
    apiKey,
    setApiKey,
    isApiKeyValid,
    selectedProvider,
    setSelectedProvider,
    selectedModel,
    setSelectedModel: handleModelSelect,
    isLoading,
    error,
    generateText,
    suggestFormula,
    analyzeData,
    generateChart,
    availableModels: AI_MODELS,
    macrosEnabled,
    toggleMacrosEnabled,
    executionHistory,
    checkMacroRequirement,
    generateVBACode,
    executeOfficeJsCode,
    executeMacro,
    processRequest,
    executeRequest,
    tokenUsageStats,
    resetTokenUsageStats,
    calculateEstimatedCost
  };

  return <AIContext.Provider value={value}>{children}</AIContext.Provider>;
};

// Custom hook to use the AI context
export const useAI = () => {
  const context = useContext(AIContext);
  if (context === undefined) {
    throw new Error('useAI must be used within an AIProvider');
  }
  return context;
}; 