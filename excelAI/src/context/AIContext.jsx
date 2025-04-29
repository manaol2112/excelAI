import React, { createContext, useState, useContext, useEffect } from 'react';
import OpenAIService from '../services/openaiService';

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
  }, []);

  // Save selected model when it changes
  useEffect(() => {
    localStorage.setItem('excelai_selected_model', JSON.stringify(selectedModel));
  }, [selectedModel]);

  // Function to handle API requests with loading state
  const callAIWithLoading = async (apiFunction, ...args) => {
    setIsLoading(true);
    setError(null);
    try {
      if (!aiService) {
        throw new Error('AI service not initialized. Please enter your API key.');
      }
      const result = await apiFunction.apply(aiService, [...args, selectedModel.id]);
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

  // Handle model selection
  const handleModelSelect = (model) => {
    setSelectedModel(model);
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