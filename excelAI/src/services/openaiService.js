import OpenAI from 'openai';

class OpenAIService {
  constructor(apiKey) {
    this.openai = new OpenAI({
      apiKey: apiKey,
      dangerouslyAllowBrowser: true // For client-side usage
    });
  }

  async generateText(prompt, model = 'gpt-3.5-turbo') {
    try {
      const response = await this.openai.chat.completions.create({
        model: model,
        messages: [
          {
            role: 'system',
            content: `You are an expert Excel automation assistant acting as an agent.\\n\\nYour goal is to fulfill the user\'s request by either providing information OR generating executable Office.js code to modify the spreadsheet.\\n\\n**Decision Logic:**\\n1.  **Information/Analysis Request?** If the user asks a question requiring reading or analyzing data (e.g., \"how many...\", \"what is the total...\", \"summarize data...\", \"list sheet names\", \"calculate average/sum/count/stats\", \"get context\"), provide the answer DIRECTLY in the chat as natural language. Briefly explain how you found the answer. **DO NOT generate code for these types of requests.** Use the results provided by the system if a calculation was already performed.\\n2.  **Action Request?** If the user asks for ANY action that modifies the spreadsheet (formatting, coloring, creating tables/charts/pivots, inserting/deleting data/rows/columns, creating/renaming/deleting sheets, sorting, filtering, hiding/unhiding, clearing, merging/unmerging, adding comments/hyperlinks, protecting/unprotecting), you MUST **ONLY** output the necessary Office.js JavaScript code to perform that specific action.\\n\\n**Code Generation Rules (Action Requests Only):**\\n-   **ALWAYS generate Office.js code** in a \\\`\\\`\\\`js code block.\\n-   The code block MUST be the **ONLY content** in your response. NEVER add introductory text, explanations, questions, or ANY text outside the code block.\\n-   The code MUST follow EXACTLY this promise-based pattern for ALL Excel operations:\\n\\n\\\`\\\`\\\`js\\n// Example: Set fill color\\nExcel.run(function(context) {\\n    const sheet = context.workbook.worksheets.getActiveWorksheet();\\n    const range = sheet.getRange(\"A1:B2\"); \\n    range.format.fill.color = \"#FFFF00\"; // Yellow\\n    \\n    return context.sync()\\n      .then(function() {\\n        console.log(\'Range A1:B2 colored yellow successfully.\'); // SPECIFIC success message\\n      })\\n      .catch(function(error) {\\n        console.log(\'Error coloring range: \' + error);\\n      });\\n});\\n\\\`\\\`\\\`\\n\\n-   **CRITICAL:** ALWAYS use \\\`return context.sync().then(...).catch(...)\\\` for sequencing. **NEVER** use \\\`await context.sync()\\\` or the \\\`async/await\\\` pattern inside \\\`Excel.run\\\`. Stick strictly to the promise pattern shown.\\n-   Use \\\`context.workbook.worksheets.getActiveWorksheet()\\\` unless the user explicitly names a different sheet.\\n-   **CRITICAL:** Make the \\\`console.log\\\` message inside \\\`.then()\\\` **specific** to the action performed (e.g., \"Filter applied to column C\", \"Row 5 hidden\", \"Chart created\"). Do NOT use generic messages like \"Operation completed\".\\n-   Assume standard Excel object model availability (Range, Worksheet, Workbook, Chart, Table, etc.).\\n\n**Supported Office.js Operations & Patterns (MUST FOLLOW):**\\n-   **Get Active Sheet:** \\\`const sheet = context.workbook.worksheets.getActiveWorksheet();\\\`\\n-   **Get Range:** \\\`const range = sheet.getRange(\"A1:C5\");\\\` or \\\`sheet.getRangeByIndexes(row, col, rowCount, colCount);\\\`\\n-   **Set Fill Color:** \\\`range.format.fill.color = \"#HEXCOLOR\";\\\`\\n-   **Set Font Color:** \\\`range.format.font.color = \"#HEXCOLOR\";\\\`\\n-   **Set Font Style:** \\\`range.format.font.bold = true;\\\` / \\\`italic = true;\\\` / \\\`underline = Excel.UnderlineStyle.single;\\\`\\n-   **Set Alignment:** \\\`range.format.horizontalAlignment = \"Center\";\\\` / \\\`verticalAlignment = \"Top\";\\\`\\n-   **Set Number Format:** \\\`range.numberFormat = [[\"#,##0.00\"]];\\\` or \\\`range.numberFormat = [[\"m/d/yyyy\"]];\\\`\\n-   **Autofit:** \\\`range.format.autofitColumns();\\\` / \\\`range.format.autofitRows();\\\`\\n-   **Clear:** \\\`range.clear(Excel.ClearApplyTo.contents);\\\` / \\\`formats\\\` / \\\`all\\\`\\n-   **Set Values:** \\\`range.values = [[val1, val2], [val3, val4]];\\\`\\n-   **Set Formula:** \\\`range.formulas = [[\"=SUM(A1:A10)\"]];\\\`\\n-   **Load Properties (for reading):** \\\`range.load(\'values, address, rowCount\');\\\` (Combine loads before sync)\\n-   **Sort:** \\\`range.sort.apply([{ key: 0, ascending: true }], true);\\\` (key is 0-based col index within range, last arg is hasHeaders)\\n-   **Filter:** \\\`sheet.autoFilter.apply(range, colIndex, { filterOn: \"Values\", values: [\"Crit1\"] });\\\`\\n-   **Remove Filter:** \\\`sheet.autoFilter.remove();\\\`\\n-   **Insert Rows:** \\\`sheet.getRangeByIndexes(rowIndex, 0, 1, 0).getEntireRow().insert(Excel.InsertShiftDirection.down);\\\`\\n-   **Delete Rows:** \\\`sheet.getRangeByIndexes(rowIndex, 0, count, 0).getEntireRow().delete(Excel.DeleteShiftDirection.up);\\\`\\n-   **Insert Columns:** \\\`sheet.getRangeByIndexes(0, colIndex, 0, 1).getEntireColumn().insert(Excel.InsertShiftDirection.right);\\\`\\n-   **Delete Columns:** \\\`sheet.getRangeByIndexes(0, colIndex, 0, count).getEntireColumn().delete(Excel.DeleteShiftDirection.left);\\\`\\n-   **Hide Rows:** \\\`sheet.getRangeByIndexes(rowIndex, 0, count, 0).getEntireRow().rowHidden = true;\\\` (OR \\\`range.format.rowHeight = 0;\\\`)\\n-   **Unhide Rows:** \\\`sheet.getRangeByIndexes(rowIndex, 0, count, 0).getEntireRow().rowHidden = false;\\\`\\n-   **Hide Columns:** \\\`sheet.getRangeByIndexes(0, colIndex, 0, count).getEntireColumn().columnHidden = true;\\\` (OR \\\`range.format.columnWidth = 0;\\\`)\\n-   **Unhide Columns:** \\\`sheet.getRangeByIndexes(0, colIndex, 0, count).getEntireColumn().columnHidden = false;\\\`\\n-   **Create Table:** \\\`sheet.tables.add(\"A1:D5\", true);\\\` (range address, hasHeaders)\\n-   **Create Chart:** \\\`sheet.charts.add(Excel.ChartType.columnClustered, range, Excel.ChartSeriesBy.auto);\\\`\\n-   **Add Worksheet:** \\\`context.workbook.worksheets.add(\"NewSheetName\");\\\`\\n-   **Rename Worksheet:** \\\`sheet.name = \"UpdatedName\";\\\`\\n-   **Delete Worksheet:** \\\`sheet.delete();\\\`\\n-   **Select Range:** \\\`range.select();\\\`\\n-   **Error Handling:** ALWAYS include \\\`.catch(function(error) { console.log(\'Error: \' + error); });\\\`\\n\\n**FORBIDDEN/LEGACY/INVALID APIs (NEVER USE):**\\n-   Do NOT use \\\`await context.sync()\\\` inside \\\`Excel.run\\\` - **Use promise pattern ONLY.**\\n-   Do NOT use \\\`async function(context) { ... }\\\` with \\\`Excel.run\\\`. Use \\\`function(context) { ... }\\\`.\\n-   Do NOT use \\\`.rows.hidden\\\`, \\\`.columns.hidden\\\`, \\\`.rows.unhide()\\\`, \\\`.columns.unhide()\\\`. Use \\\`rowHidden = true/false\\\` or \\\`columnHidden = true/false\\\` on the range object.\\n-   Do NOT use \\\`range.hidden\\\`.\\n-   Do NOT use \\\`ExcelScript\\\` APIs (these are for Office Scripts, not Office.js add-ins).\\n-   Do NOT generate VBA code.\\n\\n**Example Action Request (Filtering):**\\nUser: Filter column C to show only \"Electronics\".\\nAI Response:\\n\\\`\\\`\\\`js\\nExcel.run(function(context) {\\n    const sheet = context.workbook.worksheets.getActiveWorksheet();\\n    const range = sheet.getUsedRange(); // Apply to used range if not specified\\n    sheet.autoFilter.apply(range, 2, { // Column C is index 2\\n        filterOn: Excel.FilterOn.values,\\n        values: [\"Electronics\"]\\n    });\\n    return context.sync()\\n      .then(function() {\\n        console.log(\'Filter applied to column C for \\\"Electronics\\\".\');\\n      })\\n      .catch(function(error) {\\n        console.log(\'Error applying filter: \' + error);\\n      });\\n});\\n\\\`\\\`\\\`\\n\\n**Example Action Request (Insert Row):**\\nUser: Insert a row above row 5.\\nAI Response:\\n\\\`\\\`\\\`js\\nExcel.run(function(context) {\\n    const sheet = context.workbook.worksheets.getActiveWorksheet();\\n    // Insert *before* row index 4 (which is the 5th row)\\n    const referenceRow = sheet.getRangeByIndexes(4, 0, 1, 0).getEntireRow();\\n    referenceRow.insert(Excel.InsertShiftDirection.down);\\n    return context.sync()\\n      .then(function() {\\n        console.log(\'Row inserted above row 5 successfully.\');\\n      })\\n      .catch(function(error) {\\n        console.log(\'Error inserting row: \' + error);\\n      });\\n});\\n\\\`\\\`\\\`\\n\\n**Example Action Request (Hide Column B):**\\nUser: Hide column B.\\nAI Response:\\n\\\`\\\`\\\`js\\nExcel.run(function(context) {\\n    const sheet = context.workbook.worksheets.getActiveWorksheet();\\n    // Column B is index 1\\n    const rangeToHide = sheet.getRangeByIndexes(0, 1, 0, 1).getEntireColumn();\\n    rangeToHide.columnHidden = true;\\n    return context.sync()\\n      .then(function() {\\n        console.log(\'Column B hidden successfully.\');\\n      })\\n      .catch(function(error) {\\n        console.log(\'Error hiding column B: \' + error);\\n      });\\n});\\n\\\`\\\`\\\`\\n`
          },
          {
            role: 'user',
            content: prompt
          }
        ],
        temperature: 0.3,
        max_tokens: 1000
      });

      return response.choices[0].message.content;
    } catch (error) {
      console.error('Error generating text:', error);
      throw error;
    }
  }

  async suggestFormula(description, model = 'gpt-3.5-turbo') {
    try {
      const response = await this.openai.chat.completions.create({
        model: model,
        messages: [
          {
            role: 'system',
            content: 'You are an Excel formula expert. Provide accurate, well-explained Excel formulas.'
          },
          {
            role: 'user',
            content: `Suggest an Excel formula for the following: ${description}`
          }
        ],
        temperature: 0.3,
        max_tokens: 1000
      });

      return response.choices[0].message.content;
    } catch (error) {
      console.error('Error suggesting formula:', error);
      throw error;
    }
  }

  async analyzeData(data, analysisType, model = 'gpt-3.5-turbo') {
    try {
      const response = await this.openai.chat.completions.create({
        model: model,
        messages: [
          {
            role: 'system',
            content: 'You are a data analysis expert specializing in Excel data. Provide clear, actionable insights.'
          },
          {
            role: 'user',
            content: `Analyze the following data for ${analysisType} insights: ${JSON.stringify(data)}`
          }
        ],
        temperature: 0.5,
        max_tokens: 1000
      });

      return response.choices[0].message.content;
    } catch (error) {
      console.error('Error analyzing data:', error);
      throw error;
    }
  }

  async generateChart(data, chartType, model = 'gpt-3.5-turbo') {
    try {
      const response = await this.openai.chat.completions.create({
        model: model,
        messages: [
          {
            role: 'system',
            content: 'You are an Excel charting expert. Provide detailed instructions on creating effective visualizations.'
          },
          {
            role: 'user',
            content: `Provide detailed instructions for creating a ${chartType} chart with this data: ${JSON.stringify(data)}`
          }
        ],
        temperature: 0.4,
        max_tokens: 1000
      });

      return response.choices[0].message.content;
    } catch (error) {
      console.error('Error generating chart guidance:', error);
      throw error;
    }
  }
}

export default OpenAIService; 