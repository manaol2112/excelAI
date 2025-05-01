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
  CardHeader
} from "@fluentui/react-components";
import { useAI } from '../../context/AIContext';
import ChatMessage from './chat/ChatMessage';
import ChatInput from './chat/ChatInput';
import ChatHeader from './chat/ChatHeader';
import SuggestionsList from './chat/SuggestionsList';
import EmptyChat from './chat/EmptyChat';

// Define AI operation modes with descriptions
const AI_MODES = {
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
    boxShadow: tokens.shadow4,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground1
  },
  messagesContainer: {
    flexGrow: 1,
    overflow: "auto",
    padding: "16px",
    display: "flex",
    flexDirection: "column",
    gap: "12px"
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
  }
});

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
    suggestChart
  } = useAI();
  
  // Local state
  const [inputValue, setInputValue] = useState('');
  const [processingAction, setProcessingAction] = useState(false);
  const [aiMode, setAiMode] = useState(AI_MODES.ASK);
  const [actionText, setActionText] = useState('');
  const [excelDataLoaded, setExcelDataLoaded] = useState(false);
  const [suggestionsVisible, setSuggestionsVisible] = useState(true);
  
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
      } catch (err) {
        console.error('Error loading Excel data:', err);
        setExcelDataLoaded(false);
      }
    }
    
    loadData();
  }, [loadExcelContext]);
  
  /**
   * Process a user message and generate an AI response
   */
  const handleSendMessage = async (message = inputValue) => {
    if (!message.trim()) return;
    
    try {
      // Clear input and set processing state
      setInputValue('');
      setProcessingAction(true);
      
      // Add user message to chat
      addMessage({ role: 'user', content: message });
      
      // Add AI "thinking" message
      addMessage({ role: 'assistant', content: '...', isThinking: true });
      
      // Force refresh Excel context before processing
      if (!excelDataLoaded) {
        await loadExcelContext(true);
        setExcelDataLoaded(true);
      }
      
      // Analyze user request to determine intent
      let requestType = 'general';
      const lowerMessage = message.toLowerCase();
      
      if (lowerMessage.includes('formula') || lowerMessage.includes('calculate') || 
          lowerMessage.match(/how (to|do) (i|we|you) (calculate|compute|determine)/i)) {
        requestType = 'formula';
      } else if (lowerMessage.includes('chart') || lowerMessage.includes('graph') || 
                lowerMessage.includes('visualize') || lowerMessage.includes('plot')) {
        requestType = 'chart';
      } else if (lowerMessage.includes('analyze') || lowerMessage.includes('find patterns') || 
                lowerMessage.includes('insights') || lowerMessage.includes('trends')) {
        requestType = 'analyze';
      } else if (lowerMessage.includes('code') || lowerMessage.includes('script') || 
                lowerMessage.includes('macro') || lowerMessage.includes('automate')) {
        requestType = 'code';
      }
      
      // Generate AI response based on mode and request type
      let response;
      
      if (aiMode.key === 'AGENT' && (requestType === 'analyze' || requestType === 'chart')) {
        // For agent mode, use specialized methods for data analysis or charts
        if (requestType === 'analyze') {
          response = await analyzeData(message);
        } else if (requestType === 'chart') {
          response = await suggestChart(message);
        }
      } else if (requestType === 'formula') {
        // For formula requests, use specialized formula method
        response = await suggestFormula(message);
      } else {
        // Default to general text generation
        response = await generateText(message, requestType);
      }
      
      // Process AI response based on request type and response content
      if (response.success) {
        // Update the thinking message with the actual response
        updateLastMessage({ 
          content: response.content || response.analysis || response.suggestion || response.formula, 
          isThinking: false 
        });
        
        // For code responses, look for executable Office.js code
        if (aiMode.key === 'AGENT' && response.content && 
            (response.content.includes('```js') || response.content.includes('```javascript'))) {
          // Extract code blocks from markdown
          const codeBlocks = response.content.match(/```(?:js|javascript)\s*([\s\S]*?)```/g);
          
          if (codeBlocks && codeBlocks.length > 0) {
            // Ask user if they want to execute the code
            setActionText('Would you like me to execute this code?');
            
            // In a real implementation, you would wait for user confirmation before executing
            // For this demo, we'll assume the user wants to execute the first code block
            
            // Extract code from markdown block
            const code = codeBlocks[0].replace(/```(?:js|javascript)\s*/, '').replace(/```$/, '');
            
            // Execute the code with Office.js
            const executeResult = await executeExcelCode(code);
            
            if (executeResult.success) {
              addMessage({ 
                role: 'assistant', 
                content: 'I\'ve successfully executed the code. ' + 
                         (executeResult.message || 'The changes have been applied to your spreadsheet.'),
                isSuccess: true
              });
            } else {
              addMessage({ 
                role: 'assistant', 
                content: 'I encountered an error while trying to execute the code: ' + 
                         (executeResult.error || 'Unknown error'),
                isError: true
              });
            }
          }
        }
      } else {
        // Update the thinking message with the error
        updateLastMessage({ 
          content: `I'm sorry, but I encountered an error: ${response.error || 'Unknown error occurred'}`,
          isThinking: false,
          isError: true
        });
      }
    } catch (err) {
      console.error('Error sending message:', err);
      
      // Update the thinking message with the error
      updateLastMessage({ 
        content: `I'm sorry, but something went wrong: ${err.message || 'Unknown error occurred'}`,
        isThinking: false,
        isError: true
      });
    } finally {
      setProcessingAction(false);
      setActionText('');
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
  const handleInputChange = (_, newValue) => {
    setInputValue(newValue || '');
  };
  
  /**
   * Handle key down events in the input
   */
  const handleKeyDown = (event) => {
    if (event.key === 'Enter' && !event.shiftKey) {
      event.preventDefault();
      handleSendMessage();
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
  const handleModeChange = (mode) => {
    setAiMode(mode);
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
            mode={aiMode}
            suggestions={SUGGESTIONS[aiMode.key]}
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
            onResend={() => {
              if (message.role === 'user') {
                handleSendMessage(message.content);
              }
            }}
          />
        ))}
        <div ref={messagesEndRef} />
        
        {suggestionsVisible && messages.length > 0 && !isProcessing && (
          <SuggestionsList 
            suggestions={SUGGESTIONS[aiMode.key]} 
            onSuggestionClick={handleSuggestionClick}
          />
        )}
      </div>
    );
  };
  
  return (
    <div className={classes.root}>
      <div className={classes.chatContainer}>
        <ChatHeader 
          mode={aiMode} 
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
          disabled={isProcessing || !apiKey}
        />
        
        {(isProcessing || processingAction) && (
          <div className={classes.loadingOverlay}>
            <Spinner size="medium" label="Processing..." />
            {actionText && (
              <Text className={classes.loadingText}>{actionText}</Text>
            )}
          </div>
        )}
      </div>
    </div>
  );
} 