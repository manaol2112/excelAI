import OpenAI from 'openai';

class OpenAIService {
  constructor(apiKey) {
    this.apiKey = apiKey;
    this.setupClient(apiKey);
    this.defaultModel = 'gpt-4-turbo';
  }

  /**
   * Initialize the OpenAI client
   */
  setupClient(apiKey) {
    if (apiKey) {
    this.openai = new OpenAI({
      apiKey: apiKey,
      dangerouslyAllowBrowser: true // For client-side usage
    });
      console.log('OpenAI client initialized successfully');
    } else {
      console.warn('No API key provided, OpenAI client not fully initialized');
      this.openai = null;
    }
  }

  /**
   * Set or update the API key
   * @param {string} newKey - The new OpenAI API key
   */
  setApiKey(newKey) {
    if (!newKey) {
      throw new Error('API key is required');
    }
    
    this.apiKey = newKey;
    this.setupClient(newKey);
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
      if (!this.openai) {
        throw new Error('OpenAI client not initialized. Please set a valid API key.');
      }

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
          if (msg && msg.role && msg.content) {
            messages.push({ role: msg.role, content: msg.content });
          }
        });
      }
      
      // Add the current prompt as a user message if not already in history
      if (!conversationHistory || conversationHistory.length === 0 || 
          conversationHistory[conversationHistory.length - 1]?.role !== 'user' || 
          conversationHistory[conversationHistory.length - 1]?.content !== prompt) {
        messages.push({ role: 'user', content: prompt });
      }
      
      // Determine which model to use based on complexity
      let modelToUse = model || this.defaultModel;
      if (promptData.dataProfile || (contextInfo && contextInfo.hasComplexData)) {
        // Use the more capable model for complex data
        modelToUse = 'gpt-4-turbo';
      } else if (promptData.requestType === 'formula' || promptData.requestType === 'code') {
        // Also use better model for formula or code generation
        modelToUse = 'gpt-4-turbo';
      }

      console.log(`Using model: ${modelToUse} for request type: ${promptData.requestType || 'general'}`);

      // Make the API call using the SDK
      const response = await this.openai.chat.completions.create({
        model: modelToUse,
        messages: messages,
        temperature: this.getTemperatureForRequestType(promptData.requestType),
        max_tokens: 1500,
        top_p: 1,
        frequency_penalty: 0,
        presence_penalty: 0,
      });

      console.log('OpenAI response received:', response.choices.length > 0 ? 'success' : 'no choices returned');

      if (!response.choices || response.choices.length === 0) {
        throw new Error('No response from OpenAI API');
      }

      // Return the text content
      return {
        success: true,
        content: response.choices[0].message.content,
        model: response.model,
        usage: response.usage
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
    let systemMessage = `You are an AI assistant specializing in Microsoft Excel. You're helping a user with their Excel spreadsheet tasks.

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
      
      // Determine if a selection is actively being used for analysis
      const hasActiveSelection = excelContext.selection && 
                               excelContext.selection.values && 
                               Array.isArray(excelContext.selection.values) &&
                               excelContext.selection.values.length > 0;
      
      if (hasActiveSelection) {
        systemMessage += `\n--- ACTIVE SELECTION FOR ANALYSIS ---\n`;
        systemMessage += `Selected Range: ${excelContext.selection.address}\n`;
        systemMessage += `Dimensions: ${excelContext.selection.values.length} rows × ${excelContext.selection.values[0]?.length || 0} columns\n`;
        
        if (excelContext.selection.hasHeaders) {
          systemMessage += `Headers: Yes (first row contains column names)\n`;
        }
        
        systemMessage += `\nIMPORTANT: Use ONLY this selected range data for analysis, not the entire worksheet data.`;
        systemMessage += `\nThe user's question specifically refers to the data they've selected in this range.\n`;
        
        // Add explicit instruction to prioritize JSON data for analysis
        systemMessage += `\nCRITICAL: When analyzing data, ALWAYS prioritize using the JSON representation of the data when it's provided, as it preserves data types and structure most accurately. The JSON format should be your primary source for analysis.\n`;
      } else if (excelContext.selectedRange) {
        systemMessage += `Selected Range: ${excelContext.selectedRange}\n`;
      }
      
      if (excelContext.usedRange) {
        systemMessage += `Used Range: ${excelContext.usedRange.address || "N/A"}\n`;
      }
      
      // Add worksheets info
      if (excelContext.worksheets && excelContext.worksheets.length > 0) {
        systemMessage += `\nWorksheets: ${excelContext.worksheets.join(', ')}\n`;
      }

      // Add data overview if available, but prioritize selection data for analysis
      if (hasActiveSelection) {
        systemMessage += `\nSelection Data Sample (${excelContext.selection.address}):\n`;
        const maxRows = Math.min(5, excelContext.selection.values.length);
        for (let i = 0; i < maxRows; i++) {
          const rowIndex = i === 0 && excelContext.selection.hasHeaders ? 'Header' : i + 1;
          systemMessage += `Row ${rowIndex}: ${JSON.stringify(excelContext.selection.values[i])}\n`;
        }
        if (excelContext.selection.values.length > maxRows) {
          systemMessage += `... (${excelContext.selection.values.length - maxRows} more rows)\n`;
        }
      } else if (excelContext.selection && excelContext.selection.values) {
        systemMessage += `\nCurrent Selection Data Sample:\n`;
        const maxRows = Math.min(5, excelContext.selection.values.length);
        for (let i = 0; i < maxRows; i++) {
          systemMessage += `${JSON.stringify(excelContext.selection.values[i])}\n`;
        }
        if (excelContext.selection.values.length > maxRows) {
          systemMessage += `... (${excelContext.selection.values.length - maxRows} more rows)\n`;
        }
      }
      
      // Include sample of used range data if available - but make it clear this is not the focus for analysis
      if (!hasActiveSelection && excelContext.usedRangeValues && excelContext.usedRangeValues.length > 0) {
        systemMessage += `\nWorksheet Data Sample (not selected by user):\n`;
        const maxRows = Math.min(5, excelContext.usedRangeValues.length);
        for (let i = 0; i < maxRows; i++) {
          systemMessage += `${JSON.stringify(excelContext.usedRangeValues[i])}\n`;
        }
        if (excelContext.usedRangeValues.length > maxRows) {
          systemMessage += `... (${excelContext.usedRangeValues.length - maxRows} more rows)\n`;
        }
      }
    }

    // Add data profile if available (for data analysis)
    if (dataProfile && (requestType === 'analyze' || requestType === 'chart')) {
      systemMessage += `\n--- DATA PROFILE ---\n`;
      
      // Add basic data statistics
      systemMessage += `Data Range: ${dataProfile.range || "Not specified"}\n`;
      if (dataProfile.rowCount) systemMessage += `Rows: ${dataProfile.rowCount}, Columns: ${dataProfile.columnCount}\n`;
      if (dataProfile.hasHeaders !== undefined) systemMessage += `Has Headers: ${dataProfile.hasHeaders ? 'Yes' : 'No'}\n`;
      if (dataProfile.completeness !== undefined) systemMessage += `Completeness: ${Math.round(dataProfile.completeness * 100)}%\n\n`;
      
      // Add explicit instructions for JSON data prioritization
      systemMessage += `IMPORTANT: The data will be provided in multiple formats, including a JSON array where each object represents a row with properly named columns. ALWAYS prioritize using the JSON format for your analysis, as it preserves data types and structure most accurately.\n\n`;
      
      // Add column information if available and relevant
      if (dataProfile.columns && dataProfile.columns.length > 0) {
        systemMessage += `Columns:\n`;
        dataProfile.columns.forEach(column => {
          if (!column) return;
          
          systemMessage += `- ${column.name || "Unnamed"} (${column.dataType || "unknown"}): ${column.nonEmptyCount || 0} non-empty values`;
          
          if (column.dataType === 'numeric' && !column.empty) {
            systemMessage += `, range: ${column.min} to ${column.max}, avg: ${column.mean?.toFixed(2) || "N/A"}`;
          }
          
          systemMessage += `\n`;
        });
      }
      
      // Add insights if available
      if (dataProfile.insights && dataProfile.insights.length > 0) {
        systemMessage += `\nInsights:\n`;
        dataProfile.insights.forEach(insight => {
          if (insight && insight.message) {
            systemMessage += `- ${insight.message}\n`;
          }
        });
      }
    }

    // General instructions for all responses
    systemMessage += `\n--- RESPONSE GUIDELINES ---
1. Always provide accurate Excel information and formulas
2. When suggesting Excel formulas, ensure they are syntactically correct
3. If you don't know something, say so rather than making up information
4. For data analysis, reference specific cell values from the user's selection
5. Keep responses concise and focused on the user's question
6. Use proper Excel terminology and concepts
7. When providing code snippets for Office.js, ensure they work in Excel's JavaScript API
8. Consider the user's Excel context in your responses

For data analysis requests:
- Only analyze data that the user has selected
- If no data is selected, politely ask the user to select a range first
- Never make up or hallucinate data points
- Always verify calculations using the provided data
- PRIORITIZE using the JSON representation of the data when available for the most accurate analysis`;

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
        return `You are focused on directly analyzing Excel data and providing EXACT answers to user questions.

EXTREMELY IMPORTANT - ZERO HALLUCINATION POLICY:
- NEVER add ANY information not explicitly in the data
- EVERY claim you make MUST be directly traceable to specific data points
- You MUST decline to answer if the data doesn't contain the requested information
- DO NOT include ANY background knowledge or context not in the data
- DO NOT suggest formulas like =COUNTIF(), =SUM(), =AVERAGE() - provide the actual answer
- Directly state counts, sums, averages or other calculations based on the data you see
- When asked "How many X in column Y?", respond with "There are 7 X in column Y" not "You can use =COUNTIF(Y:Y, X)"
- When asked about maximum or minimum values, provide the actual values and their locations
- When asked about trends or patterns, analyze the data directly and explain what you find
- NEVER respond with "You can use..." or "Try using..." statements suggesting formulas or functions
- If asked for information that doesn't exist in the data, clearly state: "The data does not contain this information"

For example:
- If asked "How many closed status in column I?", answer "There are 5 Closed status items in column I."
- If asked "Which sales rep had the highest revenue?", answer "John Smith had the highest revenue at $53,200."
- If asked "What's the average sale amount?", answer "The average sale amount is $1,250."

Always answer as if you've already performed the calculation or analysis the user wants.`;

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
        return `You are a data analysis assistant for Excel. Your primary purpose is to directly answer questions about the user's data.

EXTREMELY IMPORTANT:
- ALWAYS provide direct answers to questions based on the data you can see
- DO NOT suggest Excel formulas or functions like COUNTIF, VLOOKUP, etc.
- Assume all calculations have already been done for you
- Respond with specific numbers and insights directly from the data
- Speak definitively about what the data shows, not how to find it

For example:
- If asked "How many sales in Q1?", respond with "There were 145 sales in Q1" not "You can count this by..."
- If asked "Which region has the highest average?", respond with "The East region has the highest average at 82.3" not "To find this, you would..."

Your goal is to be a direct answer engine, not a formula suggestion tool.`;
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
   * Analyze Excel data and provide insights with improved handling of selection data
   * @param {string} prompt - The analysis request
   * @param {object} context - Context including data profile and Excel context
   * @param {string} model - The model ID to use
   * @returns {Promise<object>} The analysis response
   */
  async analyzeData(prompt, context, model = 'gpt-3.5-turbo') {
    try {
      console.log("Starting data analysis with prompt:", prompt);
      
      // Check if we have selection data in the context
      const hasSelectionData = context?.excelContext?.selection?.values && 
                              Array.isArray(context.excelContext.selection.values) &&
                              context.excelContext.selection.values.length > 0;
      
      // Determine if we have sufficient data for analysis
      if (!hasSelectionData && (!context?.dataProfile || !context.dataProfile.columns)) {
        console.warn("Insufficient data for analysis - using a more cautious prompt");
        
        // Add a warning if the request requires data but none is available
        const analysisWords = ['analyze', 'calculate', 'count', 'sum', 'average', 'find', 'how many'];
        const needsDataAnalysis = analysisWords.some(word => prompt.toLowerCase().includes(word));
        
        if (needsDataAnalysis) {
          return {
            success: true,
            analysis: "I don't have enough data to perform this analysis. Please select a range in your Excel spreadsheet first, then ask your question again."
          };
        }
      }

      // Handle specific case where the prompt data already includes comprehensive data analysis
      const containsFullData = prompt.includes("===COMPLETE_DATASET_BEGIN===") && 
                              prompt.includes("===COMPLETE_DATASET_END===");
      
      // Check if the prompt contains JSON data
      const containsJsonData = prompt.includes("JSON representation of the data:") || 
                               prompt.includes("```json") ||
                               (prompt.includes("[{") && prompt.includes("}]"));
                              
      // If the prompt already contains a full dataset, use a dedicated approach to ensure accurate analysis
      if (containsFullData || containsJsonData) {
        console.log(`Detected ${containsJsonData ? 'JSON data' : 'full dataset'} in prompt - using specialized analysis approach`);
        
        // For comprehensive analysis of a full dataset, use the most capable model available
        const analysisModel = 'gpt-4-turbo';
        
        // Create enhanced system directive for data analysis with emphasis on JSON data
        const dataAnalysisSystemDirective = `
You are an Excel data analysis expert. You're examining a specific Excel data selection.

Follow these key guidelines:
1. Base your analysis ONLY on the exact data provided - never invent or guess data
2. PRIORITIZE using the JSON representation of the data when available, as it preserves data types and structure
3. The JSON format is the most accurate representation of the data - use it as your primary data source
4. For numerical answers, show your calculations and cite specific values from the JSON data
5. When counting items, use the pre-calculated metrics provided, and double-check your counts against the JSON data
6. For any statistical claims, reference specific values from the JSON dataset
7. Make sure your analysis includes ALL rows in the selection, particularly verifying the last row is included
8. If a question cannot be answered with the available data, clearly state this limitation

If the dataset includes a JSON representation, prioritize using that over other formats, as it preserves 
data types and relationships most accurately.

NEVER hallucinate data or make up information not present in the dataset.`;

        // Make direct API call to analyze the full dataset
        console.log("Calling OpenAI directly with full dataset for analysis using model:", analysisModel);
        const response = await this.openai.chat.completions.create({
          model: analysisModel,
          messages: [
            { role: 'system', content: dataAnalysisSystemDirective },
            { role: 'user', content: prompt }
          ],
          temperature: 0.0, // Use zero temperature for completely deterministic analysis
          max_tokens: 1500
        });
        
        if (!response.choices || response.choices.length === 0) {
          throw new Error('No response from OpenAI API for dataset analysis');
        }
        
        return {
          success: true,
          analysis: response.choices[0].message.content,
          usage: response.usage,
          model: response.model
        };
      }
      
      // For regular analysis requests, create enhanced system directives with JSON emphasis
      const enhancedSystemDirectives = `
Analyze the provided Excel data and give clear, accurate observations based ONLY on the data provided.
Follow these strict guidelines to ensure accuracy:

1. ALWAYS prioritize using the JSON representation of the data when available, as it preserves data types and structure
2. The JSON format is the most accurate representation of the data - use it as your primary data source
3. Only respond with analysis that is directly based on the provided data - never invent or guess data points
4. For any numerical question, first verify if the JSON data contains the required information
5. Show your calculations explicitly when answering quantitative questions
6. If a question cannot be answered with the available data, clearly state this limitation
7. For any statistical claims, cite specific values from the JSON data as evidence
8. When analyzing patterns or trends, only mention patterns that are clearly visible in the data
9. If the data selection is empty or insufficient, explain that more data is needed for analysis
10. For counting questions, always use the exact counts from the pre-calculated metrics in the JSON data
11. For questions about specific values, only mention values that exist in the JSON dataset
12. If asked to visualize or create charts, base recommendations only on the actual data structure
13. Ensure your analysis includes ALL rows in the dataset, particularly verifying the last row is included

When formatting responses:
- Use clear, concise language
- Structure complex analyses with bullet points or numbered lists
- For large datasets, summarize key findings rather than listing all data points
- Include relevant statistics (counts, averages, etc.) to support conclusions
- When appropriate, suggest follow-up analyses that could provide additional insights

NEVER make up or invent data that isn't explicitly present in the provided dataset.`;
      
      // Create prompt data with enhanced context
      const promptData = {
        prompt,
        excelContext: context.excelContext,
        dataProfile: context.dataProfile,
        requestType: 'analyze',
        conversationHistory: context.messages,
        systemDirectives: enhancedSystemDirectives
      };

      // Use a more powerful model for complex data analysis
      let analysisModel = model;
      
      // If dealing with a large or complex dataset, upgrade to a more capable model
      const hasComplexData = context?.excelContext?.selection?.values?.length > 10 || 
                           (context?.dataProfile?.columns?.length > 5);
      
      if (hasComplexData) {
        analysisModel = 'gpt-4-turbo'; // Use more powerful model for complex data
        console.log("Using more powerful model for complex data analysis");
      }
      
      // Generate the analysis using the appropriate model with lower temperature for accuracy
      console.log("Calling OpenAI for analysis with model:", analysisModel);
      const response = await this.openai.chat.completions.create({
        model: analysisModel,
        messages: [
          { 
            role: 'system', 
            content: this.buildSystemMessage(context.excelContext, 'analyze', context.dataProfile, enhancedSystemDirectives) 
          },
          // Include conversation history for context
          ...(context.messages || []).map(msg => ({
            role: msg.role,
            content: msg.content
          })),
          { role: 'user', content: prompt }
        ],
        temperature: 0.0, // Use low temperature for more deterministic/accurate responses
        max_tokens: 1500
      });
      
      console.log("Analysis response received:", response.choices.length > 0 ? "success" : "failed");

      if (!response.choices || response.choices.length === 0) {
        throw new Error('No response from OpenAI API');
      }

      return {
        success: true,
        analysis: response.choices[0].message.content,
        usage: response.usage,
        model: response.model
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