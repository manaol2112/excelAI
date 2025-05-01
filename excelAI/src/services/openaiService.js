import OpenAI from 'openai';

class OpenAIService {
  constructor(apiKey) {
    this.openai = new OpenAI({
      apiKey: apiKey,
      dangerouslyAllowBrowser: true // For client-side usage
    });
    this.apiEndpoint = 'https://api.openai.com/v1/chat/completions';
    this.defaultModel = 'gpt-4-turbo';
  }

  /**
   * Set or update the API key
   * @param {string} newKey - The new OpenAI API key
   */
  setApiKey(newKey) {
    if (!newKey) {
      throw new Error('API key is required');
    }
    
    // Update the API key in the OpenAI instance
    this.openai = new OpenAI({
      apiKey: newKey,
      dangerouslyAllowBrowser: true
    });
    
    console.log('OpenAI API key updated successfully');
  }

  /**
   * Generate text with enhanced context handling
   * @param {Object|string} promptData - Either a string prompt or an object containing prompt and context
   * @param {string} model - The model to use
   * @returns {Promise<string>} The generated text
   */
  async generateText(promptData, model = 'gpt-3.5-turbo') {
    try {
      // Determine if we have a simple prompt or a context-enhanced prompt
      let prompt, conversationHistory = [], contextInfo = null;
      
      if (typeof promptData === 'string') {
        prompt = promptData;
      } else {
        // Extract components from the enhanced prompt object
        prompt = promptData.prompt;
        conversationHistory = promptData.conversationHistory || [];
        contextInfo = promptData.excelContext || null;
      }
      
      // Build system message with enhanced context
      let systemMessage = this.buildSystemMessage(contextInfo, promptData.requestType, promptData.dataProfile, promptData.systemDirectives);
      
      // Build messages array for the API call
      const messages = [
        { role: 'system', content: systemMessage }
      ];
      
      // Add conversation history if available
      if (conversationHistory && conversationHistory.length > 0) {
        // Cap conversation history length to avoid token limits
        const maxHistoryMessages = Math.min(conversationHistory.length, 10);
        const recentHistory = conversationHistory.slice(-maxHistoryMessages);
        
        // Add each message from history
        recentHistory.forEach(msg => {
          messages.push({ role: msg.role, content: msg.content });
        });
      }
      
      // Add the current prompt as a user message if not already in history
      if (!conversationHistory || conversationHistory.length === 0 || 
          conversationHistory[conversationHistory.length - 1].role !== 'user' || 
          conversationHistory[conversationHistory.length - 1].content !== prompt) {
        messages.push({ role: 'user', content: prompt });
      }
      
      // Determine which model to use based on complexity
      let modelToUse = this.defaultModel;
      if (promptData.dataProfile || (contextInfo && contextInfo.hasComplexData)) {
        // Use the more capable model for complex data
        modelToUse = 'gpt-4-turbo';
      } else if (promptData.requestType === 'formula' || promptData.requestType === 'code') {
        // Also use better model for formula or code generation
        modelToUse = 'gpt-4-turbo';
      }

      // Make the API call
      const response = await fetch(this.apiEndpoint, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${this.openai.apiKey}`
        },
        body: JSON.stringify({
          model: modelToUse,
          messages,
          temperature: this.getTemperatureForRequestType(promptData.requestType),
          max_tokens: 2000,
          top_p: 1,
          frequency_penalty: 0,
          presence_penalty: 0,
        })
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error?.message || `API error: ${response.status}`);
      }

      const data = await response.json();
      
      if (!data.choices || data.choices.length === 0) {
        throw new Error('No response from OpenAI API');
      }

      // Return the text content
      return {
        success: true,
        content: data.choices[0].message.content,
        model: data.model,
        usage: data.usage
      };
    } catch (error) {
      console.error('Error in generateText:', error);
      return {
        success: false,
        error: error.message || 'Failed to generate text'
      };
    }
  }

  /**
   * Build a comprehensive system message with context
   */
  buildSystemMessage(excelContext, requestType, dataProfile, customDirectives) {
    let systemMessage = `You are an advanced Excel AI Assistant specialized in helping users with Microsoft Excel. 
Your goal is to provide accurate, helpful, and context-aware responses to Excel-related questions and tasks.

`;

    // Add custom directives if provided
    if (customDirectives) {
      systemMessage += `${customDirectives}\n\n`;
    }

    // Add request type specific instructions
    systemMessage += this.getRequestTypeInstructions(requestType);

    // Add Excel context information if available
    if (excelContext) {
      systemMessage += `\n--- CURRENT EXCEL CONTEXT ---\n`;
      
      if (excelContext.workbookName) {
        systemMessage += `Workbook: ${excelContext.workbookName}\n`;
      }
      
      if (excelContext.activeWorksheet) {
        systemMessage += `Active Worksheet: ${excelContext.activeWorksheet}\n`;
      }
      
      if (excelContext.selectedRange) {
        systemMessage += `Selected Range: ${excelContext.selectedRange}\n`;
      }
      
      if (excelContext.usedRange) {
        systemMessage += `Used Range: ${excelContext.usedRange}\n`;
      }
      
      // Add worksheets info
      if (excelContext.worksheets && excelContext.worksheets.length > 0) {
        systemMessage += `\nWorksheets: ${excelContext.worksheets.join(', ')}\n`;
      }

      // Add data overview if available
      if (excelContext.dataOverview) {
        systemMessage += `\nData Overview:\n${JSON.stringify(excelContext.dataOverview, null, 2)}\n`;
      }
    }

    // Add data profile if available (for data analysis)
    if (dataProfile && (requestType === 'analyze' || requestType === 'chart')) {
      systemMessage += `\n--- DATA PROFILE ---\n`;
      
      // Add basic data statistics
      systemMessage += `Data Range: ${dataProfile.range}\n`;
      systemMessage += `Rows: ${dataProfile.rowCount}, Columns: ${dataProfile.columnCount}\n`;
      systemMessage += `Has Headers: ${dataProfile.hasHeaders ? 'Yes' : 'No'}\n`;
      systemMessage += `Completeness: ${Math.round(dataProfile.completeness * 100)}%\n\n`;
      
      // Add column information
      systemMessage += `Columns:\n`;
      dataProfile.columns.forEach(column => {
        systemMessage += `- ${column.name} (${column.dataType}): ${column.nonEmptyCount} non-empty values`;
        
        if (column.dataType === 'numeric' && !column.empty) {
          systemMessage += `, range: ${column.min} to ${column.max}, avg: ${column.mean.toFixed(2)}`;
        }
        
        systemMessage += `\n`;
      });
      
      // Add insights if available
      if (dataProfile.insights && dataProfile.insights.length > 0) {
        systemMessage += `\nInsights:\n`;
        dataProfile.insights.forEach(insight => {
          systemMessage += `- ${insight.message}\n`;
        });
      }
    }

    // General instructions for all responses
    systemMessage += `\n--- RESPONSE GUIDELINES ---
1. Always provide accurate Excel information and formulas
2. When suggesting Excel formulas, ensure they are syntactically correct
3. If you don't know something, say so rather than making up information
4. Keep responses concise and focused on the user's question
5. Use proper Excel terminology and concepts
6. When providing code snippets for Office.js, ensure they work in Excel's JavaScript API
7. Consider the user's Excel context in your responses
8. For complex actions, break down steps clearly`;

    return systemMessage;
  }

  /**
   * Get specific instructions based on request type
   */
  getRequestTypeInstructions(requestType) {
    switch (requestType) {
      case 'formula':
        return `You are focused on helping the user create Excel formulas. Provide accurate, efficient formulas that solve their problem.
When suggesting formulas:
- Explain how the formula works and why you chose it
- If multiple approaches exist, mention alternatives
- Consider edge cases (errors, empty cells, etc.)
- Use modern Excel functions when appropriate
- Format complex formulas for readability`;

      case 'analyze':
        return `You are focused on helping the user analyze their Excel data. Provide clear insights and observations.
When analyzing data:
- Identify key trends, patterns, and outliers
- Suggest appropriate statistical methods
- Recommend visualization approaches
- Consider data quality issues
- Focus on the most relevant insights`;

      case 'chart':
        return `You are focused on helping the user create effective Excel charts. Provide recommendations for visualizing their data.
When suggesting charts:
- Recommend the most appropriate chart type for their data
- Explain why your suggestion is effective
- Include customization tips for clarity
- Consider data structure and relationships
- Suggest title, labels, and formatting`;

      case 'code':
        return `You are focused on helping the user with Office.js code for Excel. Provide working code examples.
When writing Office.js code:
- Ensure the code follows Excel JavaScript API best practices
- Include error handling and async/await patterns
- Structure code for readability and maintenance
- Consider performance implications
- Explain key parts of the code`;

      default:
        return `You are a general Excel AI Assistant. Respond helpfully to any Excel-related questions or tasks.`;
    }
  }

  /**
   * Adjust temperature based on request type for optimal responses
   */
  getTemperatureForRequestType(requestType) {
    switch (requestType) {
      case 'formula':
        return 0.1; // Lower temperature for more precise formula generation
      case 'code':
        return 0.1; // Lower temperature for code generation
      case 'analyze':
        return 0.3; // Slightly higher for analysis to encourage insights
      case 'chart':
        return 0.4; // Higher for chart suggestions for creativity
      default:
        return 0.7; // Default for general questions
    }
  }

  /**
   * Suggest an Excel formula based on a description
   * @param {string} description - Description of what the formula should do
   * @param {object} context - Excel context information
   * @param {string} model - The model ID to use
   * @returns {Promise<object>} The formula suggestion response
   */
  async suggestFormula(description, context, model = 'gpt-3.5-turbo') {
    try {
      // Build a prompt that focuses on formula generation
      const formulaPrompt = `I need an Excel formula that does the following: ${description}`;

      // Create prompt data with formula-specific context
      const promptData = {
        prompt: formulaPrompt,
        excelContext: context.excelContext,
        requestType: 'formula',
        conversationHistory: context.messages,
        systemDirectives: `Focus on generating an accurate Excel formula that matches the user's requirements. Provide explanations for how the formula works.`
      };

      // Generate the formula suggestion using the specified model
      const response = await this.generateText(promptData, model);

      if (!response.success) {
        throw new Error(response.error || 'Failed to generate formula suggestion');
      }

      // Extract the formula from the response
      // This is a simple extraction - in a real implementation, you might want
      // to parse the response more carefully to extract just the formula
      const formulaMatch = response.content.match(/`([^`]+)`|```([^`]+)```/);
      const formula = formulaMatch ? (formulaMatch[1] || formulaMatch[2]).trim() : response.content;

      return {
        success: true,
        formula,
        explanation: response.content,
        usage: response.usage
      };
    } catch (error) {
      console.error('Error in suggestFormula:', error);
      return {
        success: false,
        error: error.message || 'Failed to suggest formula'
      };
    }
  }

  /**
   * Analyze Excel data and provide insights
   * @param {string} prompt - The analysis request
   * @param {object} context - Context including data profile and Excel context
   * @param {string} model - The model ID to use
   * @returns {Promise<object>} The analysis response
   */
  async analyzeData(prompt, context, model = 'gpt-3.5-turbo') {
    try {
      // Create prompt data with analysis-specific context
      const promptData = {
        prompt,
        excelContext: context.excelContext,
        dataProfile: context.dataProfile,
        requestType: 'analyze',
        conversationHistory: context.messages,
        systemDirectives: `Analyze the provided Excel data and give clear, insightful observations. Focus on patterns, outliers, and notable trends.`
      };

      // Generate the analysis using the specified model
      const response = await this.generateText(promptData, model);

      if (!response.success) {
        throw new Error(response.error || 'Failed to analyze data');
      }

      return {
        success: true,
        analysis: response.content,
        usage: response.usage
      };
    } catch (error) {
      console.error('Error in analyzeData:', error);
      return {
        success: false,
        error: error.message || 'Failed to analyze data'
      };
    }
  }

  /**
   * Suggest a chart based on data
   * @param {string} prompt - The chart request
   * @param {object} context - Context including data profile and Excel context
   * @param {string} model - The model ID to use
   * @returns {Promise<object>} The chart suggestion response
   */
  async suggestChart(prompt, context, model = 'gpt-3.5-turbo') {
    try {
      // Create prompt data with chart-specific context
      const promptData = {
        prompt,
        excelContext: context.excelContext,
        dataProfile: context.dataProfile,
        requestType: 'chart',
        conversationHistory: context.messages,
        systemDirectives: `Suggest appropriate chart types for visualizing the data. Include specifics on how to create the chart in Excel.`
      };

      // Generate the chart suggestion using the specified model
      const response = await this.generateText(promptData, model);

      if (!response.success) {
        throw new Error(response.error || 'Failed to suggest chart');
      }

      // For a real implementation, you might parse this to extract specific chart settings
      // or even generate Office.js code to create the chart
      
      return {
        success: true,
        suggestion: response.content,
        usage: response.usage
      };
    } catch (error) {
      console.error('Error in suggestChart:', error);
      return {
        success: false,
        error: error.message || 'Failed to suggest chart'
      };
    }
  }
}

// Initialize and export a singleton instance
const serviceInstance = new OpenAIService(localStorage.getItem('openai_api_key') || '');
export default serviceInstance; 