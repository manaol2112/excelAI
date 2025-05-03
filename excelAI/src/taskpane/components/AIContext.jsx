import React, { useState, createContext, useContext } from 'react';
import { excelService } from '../../services/excelService';
import { analyzeData, suggestFormula, suggestChart, generateResponse } from '../../services/openaiService';

export const AIContext = createContext(null);

export const useAI = () => {
  const context = useContext(AIContext);
  if (!context) {
    throw new Error('useAI must be used within an AIProvider');
  }
  return context;
};

export const AIProvider = ({ children }) => {
  const [apiKey, setApiKey] = useState(() => {
    // Try to load from storage on first render
    if (typeof localStorage !== 'undefined') {
      return localStorage.getItem('openai_api_key') || '';
    }
    return '';
  });
  
  const [messages, setMessages] = useState([]);
  const [error, setError] = useState(null);
  const [isProcessing, setIsProcessing] = useState(false);

  const addMessage = (message) => {
    setMessages([...messages, message]);
  };

  const updateLastMessage = (message) => {
    if (messages.length > 0) {
      setMessages([...messages.slice(0, -1), message]);
    }
  };

  const clearMessages = () => {
    setMessages([]);
  };

  const generateText = async (prompt, type = 'general') => {
    try {
      setIsProcessing(true);
      const response = await generateResponse(prompt, type, apiKey);
      setIsProcessing(false);
      return response;
    } catch (error) {
      setIsProcessing(false);
      setError(`Error generating text: ${error.message}`);
      return { success: false, error: error.message };
    }
  };

  const executeExcelCode = async (code) => {
    try {
      setIsProcessing(true);
      const result = await excelService.execute(code);
      setIsProcessing(false);
      return result;
    } catch (error) {
      setIsProcessing(false);
      setError(`Error executing Excel code: ${error.message}`);
      return { success: false, error: error.message };
    }
  };

  const canUndo = async () => {
    const operationsHistory = await excelService.getOperationsHistory();
    return operationsHistory && operationsHistory.length > 0;
  };

  const undoLastOperation = async () => {
    try {
      setIsProcessing(true);
      const result = await excelService.undoLastOperation();
      setIsProcessing(false);
      return result;
    } catch (error) {
      setIsProcessing(false);
      setError(`Error undoing operation: ${error.message}`);
      return { success: false, error: error.message };
    }
  };

  const loadExcelContext = async (force = false) => {
    try {
      setIsProcessing(true);
      const result = await excelService.preloadWorkbookData(force);
      setIsProcessing(false);
      return result;
    } catch (error) {
      setIsProcessing(false);
      setError(`Error loading Excel context: ${error.message}`);
      return null;
    }
  };

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

  const countValueInColumn = async (columnLetter, targetValue, caseInsensitive = true) => {
    try {
      setIsProcessing(true);
      console.log(`Counting occurrences of "${targetValue}" in column ${columnLetter}`);
      const result = await excelService.countInColumn(columnLetter, targetValue, caseInsensitive);
      setIsProcessing(false);
      return result;
    } catch (error) {
      setIsProcessing(false);
      setError(`Error counting values: ${error.message}`);
      return { success: false, error: error.message };
    }
  };

  const analyzeColumnData = async (columnLetter) => {
    try {
      setIsProcessing(true);
      console.log(`Analyzing column ${columnLetter}`);
      const result = await excelService.analyzeColumn(columnLetter);
      setIsProcessing(false);
      return result;
    } catch (error) {
      setIsProcessing(false);
      setError(`Error analyzing column: ${error.message}`);
      return { success: false, error: error.message };
    }
  };

  const answerDataQuestion = async (question) => {
    try {
      setIsProcessing(true);
      console.log(`Attempting to directly answer: "${question}"`);
      
      // First, ensure we have the latest Excel context
      await loadExcelContext(true);
      
      // Identify if this is a column-based question
      const columnMatch = question.match(/column\s+([a-z])/i);
      let targetColumn = null;
      
      if (columnMatch && columnMatch[1]) {
        targetColumn = columnMatch[1].toUpperCase();
      }
      
      if (!targetColumn) {
        // Try to extract column letter from questions like "how many closed in I?"
        const columnLetterMatch = question.match(/\s([a-z])\??$/i);
        if (columnLetterMatch && columnLetterMatch[1]) {
          targetColumn = columnLetterMatch[1].toUpperCase();
        }
      }
      
      console.log(`Extracted column: ${targetColumn || 'None'}`);
      
      // Check for counting questions
      const countingQuestion = 
        question.toLowerCase().includes('how many') || 
        question.toLowerCase().includes('count') ||
        question.toLowerCase().includes('number of');
      
      if (countingQuestion && targetColumn) {
        // Extract value to count
        const commonValues = ['open', 'closed', 'pending', 'completed', 'yes', 'no', 'true', 'false'];
        let targetValue = null;
        
        for (const value of commonValues) {
          if (question.toLowerCase().includes(value)) {
            // Format properly with first letter capitalized
            targetValue = value.charAt(0).toUpperCase() + value.slice(1).toLowerCase();
            break;
          }
        }
        
        console.log(`Extracted value to count: ${targetValue || 'None'}`);
        
        if (targetValue) {
          const countResult = await countValueInColumn(targetColumn, targetValue, true);
          
          if (countResult && countResult.success) {
            const answer = `There ${countResult.count === 1 ? 'is' : 'are'} ${countResult.count} "${targetValue}" ${countResult.count === 1 ? 'entry' : 'entries'} in column ${targetColumn}.`;
            
            setIsProcessing(false);
            return { success: true, answer };
          }
        }
      }
      
      // Handle column analysis questions
      const analysisQuestion = 
        question.toLowerCase().includes('analyze') || 
        question.toLowerCase().includes('tell me about') ||
        question.toLowerCase().includes('what\'s in') ||
        question.toLowerCase().includes('describe');
      
      if (analysisQuestion && targetColumn) {
        const analysisResult = await analyzeColumnData(targetColumn);
        
        if (analysisResult && analysisResult.success && analysisResult.analysis) {
          const analysis = analysisResult.analysis;
          
          let answer = `Column ${targetColumn} contains ${analysis.nonEmptyCells} non-empty values out of ${analysis.totalCells} total cells.`;
          
          // Add top values
          if (analysis.topValues && analysis.topValues.length > 0) {
            answer += `\n\nThe most common values in this column are:`;
            analysis.topValues.slice(0, 5).forEach((item, index) => {
              answer += `\n${index + 1}. "${item.value}" appears ${item.count} time${item.count === 1 ? '' : 's'}`;
            });
          }
          
          // Add numerical stats if available
          if (analysis.numericalStats) {
            const stats = analysis.numericalStats;
            answer += `\n\nNumerical statistics:`;
            answer += `\n- Average: ${stats.average.toFixed(2)}`;
            answer += `\n- Sum: ${stats.sum}`;
            answer += `\n- Min: ${stats.min}`;
            answer += `\n- Max: ${stats.max}`;
          }
          
          setIsProcessing(false);
          return { success: true, answer };
        }
      }
      
      // If we couldn't directly answer, pass through to OpenAI
      setIsProcessing(false);
      return { success: false, error: "Could not directly answer question" };
    } catch (error) {
      setIsProcessing(false);
      setError(`Error answering question: ${error.message}`);
      return { success: false, error: error.message };
    }
  };

  const value = {
    apiKey,
    setApiKey,
    messages,
    error,
    isProcessing,
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
    countValueInColumn,
    analyzeColumnData,
    answerDataQuestion
  };

  return (
    <AIContext.Provider value={value}>
      {children}
    </AIContext.Provider>
  );
}; 