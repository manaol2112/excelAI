import OpenAI from 'openai';

class OpenAIService {
  constructor(apiKey) {
    this.openai = new OpenAI({
      apiKey: apiKey,
      dangerouslyAllowBrowser: true // For client-side usage
    });
    
    // Store token usage statistics for optimization
    this.tokenUsageStats = {
      totalRequests: 0,
      totalTokensUsed: 0,
      tokensByComplexity: {
        simple: { requests: 0, totalTokens: 0 },
        medium: { requests: 0, totalTokens: 0 },
        complex: { requests: 0, totalTokens: 0 }
      }
    };
  }

  /**
   * Estimates the token limit needed for a task based on its complexity
   * @param {string} task - The task description
   * @returns {Object} - Object containing token limit and complexity assessment
   */
  estimateTokenLimit(task) {
    // Default result
    const result = {
      tokenLimit: 1500,
      complexity: 'medium',
      explanation: 'Default medium complexity allocation'
    };
    
    // Convert task to lowercase for case-insensitive matching
    const taskLower = task.toLowerCase();
    
    // Simple operations (formatting, basic cell operations)
    if (/\b(format|color|fill|bold|italic|underline|font|align|hide|unhide|insert text|change text|set value)\b/i.test(taskLower)) {
      // Check if there might be multiple operations in one request
      const operationMatches = taskLower.match(/\b(format|color|fill|bold|italic|underline|font|align|hide|unhide|insert|change|set)\b/gi);
      
      if (operationMatches && operationMatches.length <= 2) {
        result.tokenLimit = 800;
        result.complexity = 'simple';
        result.explanation = 'Simple formatting or cell operation';
      } else {
        // Multiple simple operations
        result.tokenLimit = 1200;
        result.complexity = 'medium';
        result.explanation = 'Multiple simple operations';
      }
    }
    
    // Medium complexity operations
    else if (/\b(chart|graph|table|sort|filter|sum|average|count|formula|autofit|merge|data validation)\b/i.test(taskLower)) {
      result.tokenLimit = 1500;
      result.complexity = 'medium';
      result.explanation = 'Medium complexity operation (charts, tables, formulas)';
    }
    
    // Complex operations
    else if (/\b(pivot|pivottable|conditional formatting|conditional format|slicer|macro|script|multiple sheets|across sheets|protection|named range)\b/i.test(taskLower)) {
      result.tokenLimit = 3000;
      result.complexity = 'complex';
      result.explanation = 'Complex operation (PivotTables, conditional formatting, multi-step)';
    }
    
    // Additional token adjustments based on specific patterns
    
    // Tasks with precise ranges often require more tokens for explicit cell addressing
    if (/[A-Z]+\d+:[A-Z]+\d+/.test(task)) {
      result.tokenLimit += 200;
      result.explanation += ' + range specification';
    }
    
    // Tasks with multiple steps
    if (/\band\b|\bthen\b|\bafter\b|\bfollowed by\b/i.test(taskLower)) {
      result.tokenLimit += 500;
      result.explanation += ' + multi-step operation';
    }
    
    // Tasks that mention specific formatting like colors, styles, etc. need more tokens
    if (/\bstyle\b|\bborder\b|\bpattern\b|\bgradient\b/i.test(taskLower)) {
      result.tokenLimit += 300;
      result.explanation += ' + detailed styling';
    }
    
    // Cap the maximum tokens to avoid excessive allocation
    const MAX_TOKENS = 4000;
    if (result.tokenLimit > MAX_TOKENS) {
      result.tokenLimit = MAX_TOKENS;
      result.explanation += ' (capped at maximum)';
    }
    
    return result;
  }

  /**
   * Tracks token usage for optimization
   * @param {string} complexity - The complexity category
   * @param {number} tokensUsed - The number of tokens used
   */
  trackTokenUsage(complexity, tokensUsed) {
    this.tokenUsageStats.totalRequests++;
    this.tokenUsageStats.totalTokensUsed += tokensUsed;
    
    if (this.tokenUsageStats.tokensByComplexity[complexity]) {
      this.tokenUsageStats.tokensByComplexity[complexity].requests++;
      this.tokenUsageStats.tokensByComplexity[complexity].totalTokens += tokensUsed;
    }
    
    // Calculate averages for reporting
    const avgTokensPerRequest = Math.round(this.tokenUsageStats.totalTokensUsed / this.tokenUsageStats.totalRequests);
    console.log(`Token usage: ${tokensUsed} (${complexity}). Average: ${avgTokensPerRequest}`);
  }

  /**
   * Gets token usage statistics
   * @returns {Object} - Token usage statistics
   */
  getTokenUsageStats() {
    const stats = {...this.tokenUsageStats};
    
    // Add averages
    if (stats.totalRequests > 0) {
      stats.averageTokensPerRequest = Math.round(stats.totalTokensUsed / stats.totalRequests);
      
      Object.keys(stats.tokensByComplexity).forEach(complexity => {
        const complexityStats = stats.tokensByComplexity[complexity];
        if (complexityStats.requests > 0) {
          complexityStats.averageTokens = Math.round(complexityStats.totalTokens / complexityStats.requests);
        }
      });
    }
    
    return stats;
  }

  /**
   * Resets token usage statistics
   */
  resetTokenUsageStats() {
    this.tokenUsageStats = {
      totalRequests: 0,
      totalTokensUsed: 0,
      tokensByComplexity: {
        simple: { requests: 0, totalTokens: 0 },
        medium: { requests: 0, totalTokens: 0 },
        complex: { requests: 0, totalTokens: 0 }
      }
    };
  }

  async generateText(prompt, model = 'gpt-3.5-turbo') {
    try {
      // Estimate the tokens needed based on task complexity
      const { tokenLimit, complexity } = this.estimateTokenLimit(prompt);
      
      const response = await this.openai.chat.completions.create({
        model: model,
        messages: [
          {
            role: 'system',
            content: `You are an expert Excel automation assistant acting as an agent.\n\nYour goal is to fulfill the user\'s request by either providing information OR generating executable Office.js code to modify the spreadsheet.\n\n**Decision Logic:**\n1.  **Information/Analysis Request?** If the user asks a question requiring reading or analyzing data (e.g., \"how many items...\", \"what is the total...\", \"summarize data...\"), provide the answer DIRECTLY in the chat as natural language. Briefly explain how you found the answer. **DO NOT generate code for these.**\n2.  **Action Request?** If the user asks for ANY action that modifies the spreadsheet (formatting, coloring, creating tables/charts/pivots, inserting/deleting data/rows/columns, creating/renaming sheets, sorting, filtering, hiding/unhiding, etc.), you MUST ONLY output the necessary Office.js JavaScript code to perform that specific action. Do not include ANY other text, explanations, or questions in your response when generating code.\n\n**Code Generation Rules (Action Requests Only):**\n-   **ALWAYS generate Office.js code** in a \\u0060\\u0060\\u0060js code block.\n-   The code block MUST be the **ONLY content** in your response. NEVER add introductory text, explanations, questions, or ANY text outside the code block.\n-   The code MUST follow EXACTLY this pattern for ALL Excel operations:\n\n\\u0060\\u0060\\u0060js\nExcel.run(function(context) {\n  const sheet = context.workbook.worksheets.getActiveWorksheet();\n  // Operation code here\n  \n  return context.sync()\n    .then(function() {\n      console.log('Operation completed successfully');\n    })\n    .catch(function(error) {\n      console.log('Error: ' + error);\n    });\n});\n\\u0060\\u0060\\u0060\n\n-   IMPORTANT: ALWAYS use \`return context.sync()\` followed by \`.then()\` and \`.catch()\` promise handling, NOT \`await context.sync()\`\n-   NEVER use async/await pattern inside Excel.run - use the promise pattern shown above\n-   Use \`context.workbook.worksheets.getActiveWorksheet()\` unless the user explicitly names a different sheet\n-   IMPORTANT: In the success callback message, always describe specifically what was done (e.g., \"Filter applied to column H\" rather than generic \"Operation completed successfully\")\n\n**Office.js Best Practices and Patterns (MUST FOLLOW):**\n- To **hide a row**: set \`range.format.rowHeight = 0;\`\n- To **unhide a row**: set \`range.format.rowHeight = 15;\` (or another positive value; do NOT use \`.rows.unhide()\` or \`.hidden = false\`)\n- To **hide a column**: set \`range.format.columnWidth = 0;\`\n- To **unhide a column**: set \`range.format.columnWidth = 8.43;\` (or another positive value; do NOT use \`.columns.unhide()\` or \`.hidden = false\`)\n- To **set font color**: \`range.format.font.color = \"#0000FF\";\` (always use hex colors)\n- To **set fill color**: \`range.format.fill.color = \"#90EE90\";\`\n- To **apply filters**: \`sheet.autoFilter.apply(range, columnIndex, filterCriteria);\`\n- To **remove filters**: \`sheet.autoFilter.remove();\`\n- To **sort**: \`range.sort.apply([{ key: sortRange, ascending: true }]);\`\n- To **create a chart**: \`sheet.charts.add(Excel.ChartType.pie, range);\` (or other chart types)\n- To **select a range**: \`range.select();\`\n- To **create a table**: \`sheet.tables.add(range.getAddress(), true);\`\n- To **rename a worksheet**: \`sheet.name = \"NewName\";\`\n- To **add a worksheet**: \`context.workbook.worksheets.add(\"SheetName\");\`\n- To **delete a worksheet**: \`sheet.delete();\`\n- To **autofit columns**: \`range.format.autofitColumns();\`\n- To **autofit rows**: \`range.format.autofitRows();\`\n- To **bold text**: \`range.format.font.bold = true;\`\n- To **italicize text**: \`range.format.font.italic = true;\`\n- To **underline text**: \`range.format.font.underline = Excel.UnderlineStyle.single;\`\n- To **align text**: \`range.format.horizontalAlignment = \"Center\";\`\n- To **set number format**: \`range.numberFormat = [[\"#,##0.00\"]];\`\n- To **clear contents**: \`range.clear(Excel.ClearApplyTo.contents);\`\n- To **set formula**: \`range.formulas = [[\"=SUM(A1:A10)\"]];\`\n- To **read values**: always use \`range.load('values');\` and \`await context.sync();\` before accessing \`range.values\`\n- To **read address**: \`range.load('address');\`\n- To **read row/column count**: \`range.load(['rowCount', 'columnCount']);\`\n- To **handle errors**: always use .catch and log errors\n- To **create a PivotTable**: use \`sheet.pivotTables.add(name, sourceData, destination);\`\n- To **add conditional formatting**: use \`range.conditionalFormats.add(type);\`\n- To **add data validation**: use \`range.dataValidation.rule = {...};\`\n- To **merge cells**: use \`range.merge();\`\n- To **unmerge cells**: use \`range.unmerge();\`\n- To **add a comment**: use \`sheet.comments.add(range, text);\`\n- To **protect a worksheet**: use \`sheet.protection.protect(options);\`\n- To **unprotect a worksheet**: use \`sheet.protection.unprotect();\`\n- To **create a named range**: use \`context.workbook.names.add(name, range);\`\n- To **add a slicer**: use \`sheet.slicers.add(table, column);\`\n- To **add a hyperlink**: use \`range.hyperlink = { address: url, textToDisplay: text };\`\n- To **format cells as table**: use detailed formatting instead of built-in styles for better control\n\n**FORBIDDEN/LEGACY/INVALID APIs (NEVER USE):**\n- Do NOT use \`.rows.unhide()\`, \`.columns.unhide()\`, \`.hidden = false\`, or any legacy Excel VBA patterns\n- Do NOT use \`await context.sync()\` inside Excel.run\n- Do NOT use \`range.hidden\`\n- Do NOT use \`range.rows\` or \`range.columns\` for hiding/unhiding\n- Do NOT use \`ExcelScript\` APIs (these are for Office Scripts, not Office.js add-ins)\n\n**Advanced Fallback Detection:**\nIf you determine that the user's request CANNOT be implemented with Office.js (e.g., printing, saving files locally, modifying Excel settings, etc.), CLEARLY state this limitation and suggest using VBA/macros instead. Example: \"This operation requires VBA macros and cannot be implemented with Office.js. You would need to use Excel's built-in macro functionality.\"\n\n**Example Information Request:**\nUser: How many items are under Electronics in A1:L11?\nAI Response: There are 5 items under Electronics in the range A1:L11. I found this by checking the values in column E.\n\n**Example Action Request (Formatting):**\nUser: Color the table in A1:L11 green.\nAI Response:\n\\u0060\\u0060\\u0060js\nExcel.run(function(context) {\n  const sheet = context.workbook.worksheets.getActiveWorksheet();\n  const range = sheet.getRange(\"A1:L11\");\n  range.format.fill.color = \"#90EE90\"; // light green\n  \n  return context.sync()\n    .then(function() {\n      console.log('Table cells A1:L11 colored green successfully');\n    })\n    .catch(function(error) {\n      console.log('Error: ' + error);\n    });\n});\n\\u0060\\u0060\\u0060\n\n**Example Action Request (Unhide Row):**\nUser: Unhide row 4.\nAI Response:\n\\u0060\\u0060\\u0060js\nExcel.run(function(context) {\n  const sheet = context.workbook.worksheets.getActiveWorksheet();\n  const range = sheet.getRange(\"4:4\");\n  range.format.rowHeight = 15; // Unhide by setting to default height\n  return context.sync()\n    .then(function() {\n      console.log('Row 4 unhidden successfully');\n    })\n    .catch(function(error) {\n      console.log('Error: ' + error);\n    });\n});\n\\u0060\\u0060\\u0060\n\n**Example Action Request (Hide Column):**\nUser: Hide column B.\nAI Response:\n\\u0060\\u0060\\u0060js\nExcel.run(function(context) {\n  const sheet = context.workbook.worksheets.getActiveWorksheet();\n  const range = sheet.getRange(\"B:B\");\n  range.format.columnWidth = 0; // Hide column\n  return context.sync()\n    .then(function() {\n      console.log('Column B hidden successfully');\n    })\n    .catch(function(error) {\n      console.log('Error: ' + error);\n    });\n});\n\\u0060\\u0060\\u0060\n\n**Example Action Request (Create Chart):**\nUser: Create a pie chart from A1:B5.\nAI Response:\n\\u0060\\u0060\\u0060js\nExcel.run(function(context) {\n  const sheet = context.workbook.worksheets.getActiveWorksheet();\n  const range = sheet.getRange(\"A1:B5\");\n  const chart = sheet.charts.add(Excel.ChartType.pie, range);\n  chart.setPosition(\"A7\", \"F20\"); // Position the chart below the data\n  chart.title.text = \"Data Chart\"; // Add a title\n  chart.legend.position = \"right\"; // Position legend\n  chart.height = 150;\n  chart.width = 300;\n  return context.sync()\n    .then(function() {\n      console.log('Pie chart created successfully from A1:B5');\n    })\n    .catch(function(error) {\n      console.log('Error: ' + error);\n    });\n});\n\\u0060\\u0060\\u0060\n\n**Example Action Request (Create PivotTable):**\nUser: Create a pivot table from data in A1:E20 with Product in rows and Region in columns, summarizing Sales.\nAI Response:\n\\u0060\\u0060\\u0060js\nExcel.run(function(context) {\n  const sheet = context.workbook.worksheets.getActiveWorksheet();\n  const dataRange = sheet.getRange(\"A1:E20\");\n  const destination = sheet.getRange(\"G1\");\n  \n  // Create the PivotTable\n  const pivotTable = sheet.pivotTables.add(\"SalesPivot\", dataRange, destination);\n  \n  // Load the hierarchies\n  pivotTable.hierarchies.load(\"items\");\n  \n  return context.sync()\n    .then(function() {\n      // Add fields to the PivotTable\n      // Assuming Product is in column A (index 0) and Region in column B (index 1), Sales in column C (index 2)\n      const productHierarchy = pivotTable.hierarchies.getItem(0);\n      const regionHierarchy = pivotTable.hierarchies.getItem(1);\n      const salesHierarchy = pivotTable.hierarchies.getItem(2);\n      \n      // Set Product as rows\n      pivotTable.rowHierarchies.add(productHierarchy);\n      \n      // Set Region as columns\n      pivotTable.columnHierarchies.add(regionHierarchy);\n      \n      // Set Sales as values\n      const dataHierarchy = pivotTable.dataHierarchies.add(salesHierarchy);\n      dataHierarchy.summarizeBy = \"Sum\";\n      \n      return context.sync();\n    })\n    .then(function() {\n      console.log('PivotTable created successfully with Product in rows, Region in columns, and Sales summarized');\n    })\n    .catch(function(error) {\n      console.log('Error: ' + error);\n    });\n});\n\\u0060\\u0060\\u0060`
          },
          {
            role: 'user',
            content: prompt
          }
        ],
        temperature: 0.3,
        max_tokens: tokenLimit, // Dynamically set based on complexity
        top_p: 0.95
      });

      // Track token usage for optimization if usage data is available
      if (response.usage) {
        this.trackTokenUsage(complexity, response.usage.completion_tokens);
      }

      return response.choices[0].message.content;
    } catch (error) {
      console.error('Error generating text:', error);
      // Improved error handling with more specific messages
      if (error.status === 429) {
        return "Rate limit exceeded. Please try again after a short wait.";
      } else if (error.status >= 500) {
        return "OpenAI service is currently experiencing issues. Please try again later.";
      } else if (error.code === 'insufficient_quota') {
        return "Your OpenAI API quota has been exceeded. Please check your billing details.";
      } else if (error.code === 'context_length_exceeded') {
        return "The request is too large. Try simplifying your task or breaking it into smaller parts.";
      } else {
        return `Error generating response: ${error.message || "Unknown error"}`;
      }
    }
  }

  async suggestFormula(description, model = 'gpt-3.5-turbo') {
    try {
      // Formula suggestions typically need moderate token counts
      const tokenLimit = 1200;
      
      const response = await this.openai.chat.completions.create({
        model: model,
        messages: [
          {
            role: 'system',
            content: `You are an Excel formula expert. Provide accurate, well-explained Excel formulas.

When suggesting formulas:
1. Always provide the complete, correctly formatted formula ready to use in Excel
2. Explain the formula components and how they work together
3. Consider edge cases and potential errors
4. For complex tasks, offer both simple and advanced solutions
5. Provide examples of the formula with sample data when helpful
6. Mention any Excel version limitations if relevant

The formula response should be structured as follows:
- **Formula:** [The complete formula, ready to copy and paste]
- **Explanation:** [Brief explanation of how the formula works]
- **Example:** [Optional: Show how the formula would work with sample data]
- **Tips:** [Optional: Any important considerations, limitations, or alternatives]

Remember to format array formulas properly with CTRL+SHIFT+ENTER notation or @ references for Excel 365 dynamic arrays where appropriate.`
          },
          {
            role: 'user',
            content: `Suggest an Excel formula for the following: ${description}`
          }
        ],
        temperature: 0.3,
        max_tokens: tokenLimit,
        top_p: 0.9
      });

      // Track token usage if available
      if (response.usage) {
        this.trackTokenUsage('medium', response.usage.completion_tokens);
      }

      return response.choices[0].message.content;
    } catch (error) {
      console.error('Error suggesting formula:', error);
      if (error.status === 429) {
        return "Rate limit exceeded. Please try again after a short wait.";
      } else {
        return `Error suggesting formula: ${error.message || "Unknown error"}`;
      }
    }
  }

  async analyzeData(data, analysisType, model = 'gpt-3.5-turbo') {
    try {
      // Analysis typically requires more tokens
      const tokenLimit = 2000;
      
      // Check if data is valid
      if (!data || (Array.isArray(data) && data.length === 0)) {
        return "No data provided for analysis. Please provide valid data.";
      }
      
      // Ensure data doesn't exceed token limits
      let dataString;
      try {
        dataString = JSON.stringify(data);
        // If data is too large, summarize or truncate it
        if (dataString.length > 15000) {
          console.warn("Data is too large for analysis, truncating...");
          if (Array.isArray(data)) {
            const truncatedData = data.slice(0, 50); // Reduced from 100 to 50 items
            dataString = JSON.stringify(truncatedData) + " [Data truncated due to size]";
          } else {
            dataString = dataString.substring(0, 10000) + "... [Data truncated due to size]"; // Reduced from 15000 to 10000
          }
        }
      } catch (e) {
        console.error("Error stringifying data:", e);
        return "Error processing data for analysis. The data structure may be too complex.";
      }

      const response = await this.openai.chat.completions.create({
        model: model,
        messages: [
          {
            role: 'system',
            content: `You are a data analysis expert specializing in Excel data. Provide clear, actionable insights from the data.

For ${analysisType} analysis, focus on:
1. Key statistics and metrics relevant to this analysis type
2. Notable patterns, trends, or anomalies
3. Clear visualizations recommendations for this data
4. Actionable recommendations based on the insights
5. Any limitations or caveats about the analysis

Structure your response with clear headings and bullet points where appropriate for readability.
Make your insights specific and relevant to the data provided.`
          },
          {
            role: 'user',
            content: `Analyze the following data for ${analysisType} insights: ${dataString}`
          }
        ],
        temperature: 0.5,
        max_tokens: tokenLimit,
        top_p: 0.9
      });

      // Track token usage if available
      if (response.usage) {
        this.trackTokenUsage('complex', response.usage.completion_tokens);
      }

      return response.choices[0].message.content;
    } catch (error) {
      console.error('Error analyzing data:', error);
      if (error.status === 429) {
        return "Rate limit exceeded. Please try again after a short wait.";
      } else {
        return `Error analyzing data: ${error.message || "Unknown error"}`;
      }
    }
  }

  async generateChart(data, chartType, model = 'gpt-3.5-turbo') {
    try {
      // Chart generation typically requires moderate token counts
      const tokenLimit = 1800;
      
      // Validate input data
      if (!data || (Array.isArray(data) && data.length === 0)) {
        return "No data provided for chart generation. Please provide valid data.";
      }
      
      // Process data to ensure it doesn't exceed token limits
      let dataString;
      try {
        dataString = JSON.stringify(data);
        if (dataString.length > 8000) { // Reduced from 12000 to 8000
          console.warn("Data is too large for chart generation, truncating...");
          if (Array.isArray(data)) {
            const truncatedData = data.slice(0, 30); // Reduced from 50 to 30
            dataString = JSON.stringify(truncatedData) + " [Data truncated due to size]";
          } else {
            dataString = dataString.substring(0, 8000) + "... [Data truncated due to size]";
          }
        }
      } catch (e) {
        console.error("Error stringifying data:", e);
        return "Error processing data for chart generation. The data structure may be too complex.";
      }

      const response = await this.openai.chat.completions.create({
        model: model,
        messages: [
          {
            role: 'system',
            content: `You are an Excel charting expert. Generate Office.js code to create a ${chartType} chart based on the provided data.

Your response should include Office.js code to create the ${chartType} chart with proper formatting and labels.

The code should follow the standard Office.js pattern:

\`\`\`js
Excel.run(function(context) {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  
  // Chart creation code here
  
  return context.sync()
    .then(function() {
      console.log('Chart created successfully');
    })
    .catch(function(error) {
      console.log('Error: ' + error);
    });
});
\`\`\`

Remember to include proper chart formatting, title, legend, and axis labels in your code.`
          },
          {
            role: 'user',
            content: `Generate Office.js code to create a ${chartType} chart with this data: ${dataString}`
          }
        ],
        temperature: 0.4,
        max_tokens: tokenLimit,
        top_p: 0.9
      });

      // Track token usage if available
      if (response.usage) {
        this.trackTokenUsage('medium', response.usage.completion_tokens);
      }

      return response.choices[0].message.content;
    } catch (error) {
      console.error('Error generating chart guidance:', error);
      if (error.status === 429) {
        return "Rate limit exceeded. Please try again after a short wait.";
      } else {
        return `Error generating chart guidance: ${error.message || "Unknown error"}`;
      }
    }
  }

  /**
   * Detects if a task requires VBA macro fallback
   * @param {string} task - The task description
   * @returns {Promise<Object>} - A promise that resolves with the detection result
   */
  async detectMacroRequirement(task, model = 'gpt-3.5-turbo') {
    try {
      // Macro detection requires minimal tokens
      const tokenLimit = 400;
      
      const response = await this.openai.chat.completions.create({
        model: model,
        messages: [
          {
            role: 'system',
            content: `You are an Excel automation expert. Your task is to analyze a user request and determine if it can be accomplished with Office.js (Excel JavaScript API) or if it requires VBA macros.

Office.js Capabilities:
- Reading/writing cell values and formulas
- Formatting cells (font, colors, borders, alignment)
- Creating/managing tables, charts, PivotTables
- Adding/removing worksheets, rows, columns
- Creating/modifying named ranges
- Applying filters, sorts, data validation
- Adding comments, hyperlinks
- Protecting/unprotecting worksheets
- Working with ranges, cells, and worksheets

VBA Macro Requirements (NOT possible with Office.js):
- File operations (save as, print, export to PDF/CSV)
- Accessing Windows File System or Registry
- Interacting with other applications
- Creating custom UI elements in Excel
- Accessing/modifying Excel application settings
- Advanced worksheet protection features
- Accessing external data sources directly
- Printing with specific configurations
- Automation across multiple Office applications

Analyze the following task and determine:
1. Can it be accomplished with Office.js? (Yes/No)
2. If No, explain why VBA is required and which specific Office.js limitation prevents this task
3. If Yes, explain how Office.js could accomplish this (high-level approach only)`
          },
          {
            role: 'user',
            content: `Determine if this Excel task requires VBA macros or can be done with Office.js: "${task}"`
          }
        ],
        temperature: 0.1,
        max_tokens: tokenLimit
      });

      const result = response.choices[0].message.content;
      
      // Parse the response to determine if macros are required
      const requiresMacro = 
        result.toLowerCase().includes("vba is required") || 
        result.toLowerCase().includes("requires vba") || 
        result.toLowerCase().includes("can't be accomplished with office.js") ||
        result.toLowerCase().includes("cannot be accomplished with office.js");
      
      // Track token usage if available
      if (response.usage) {
        this.trackTokenUsage('simple', response.usage.completion_tokens);
      }
      
      return {
        requiresMacro,
        explanation: result
      };
    } catch (error) {
      console.error('Error detecting macro requirement:', error);
      return {
        requiresMacro: false, // Default to false if we can't determine
        explanation: `Error detecting if macros are required: ${error.message || "Unknown error"}`
      };
    }
  }

  /**
   * Generates VBA code for tasks that cannot be accomplished with Office.js
   * @param {string} task - The task description
   * @returns {Promise<string>} - A promise that resolves with the VBA code
   */
  async generateVBACode(task, model = 'gpt-3.5-turbo') {
    try {
      // VBA code generation requires more tokens
      const tokenLimit = 2000;
      
      const response = await this.openai.chat.completions.create({
        model: model,
        messages: [
          {
            role: 'system',
            content: `You are an Excel VBA expert. Generate VBA code for tasks that cannot be accomplished with Office.js.

The VBA code should:
1. Be well-commented and explain key sections
2. Follow best practices for error handling with On Error statements
3. Use Option Explicit where appropriate
4. Be optimized for performance when possible
5. Be compatible with modern Excel versions

Structure your response as follows:
1. Brief explanation of the approach
2. Complete, ready-to-use VBA code in a code block
3. Instructions on how to add and run this macro in Excel

For tasks involving file operations, user interface modifications, or Excel application settings, ensure the VBA provides proper user feedback and error handling.`
          },
          {
            role: 'user',
            content: `Generate VBA code for the following Excel task that cannot be done with Office.js: "${task}"`
          }
        ],
        temperature: 0.2,
        max_tokens: tokenLimit,
        top_p: 0.9
      });

      // Track token usage if available
      if (response.usage) {
        this.trackTokenUsage('complex', response.usage.completion_tokens);
      }

      return response.choices[0].message.content;
    } catch (error) {
      console.error('Error generating VBA code:', error);
      if (error.status === 429) {
        return "Rate limit exceeded. Please try again after a short wait.";
      } else {
        return `Error generating VBA code: ${error.message || "Unknown error"}`;
      }
    }
  }
}

export default OpenAIService; 