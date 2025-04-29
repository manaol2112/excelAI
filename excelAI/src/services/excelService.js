class ExcelService {
  constructor() {
    // Initialize operations history array to track all operations for undo
    this.operationsHistory = [];
    this.maxHistoryLength = 50; // Limit history to prevent memory issues
    console.log("ExcelService initialized with empty operations history");
  }

  /**
   * Tracks an operation in history for potential undo
   * @param {string} type - The type of operation (e.g., "format", "insertText")
   * @param {Object} details - Details needed to undo the operation
   * @private
   */
  _trackOperation(type, details) {
    if (!this.operationsHistory) {
      this.operationsHistory = []; // Initialize if undefined
    }
    
    // Add operation to history
    this.operationsHistory.unshift({
      type,
      details,
      timestamp: new Date().getTime()
    });
    
    // Trim history if it exceeds max length
    if (this.operationsHistory.length > this.maxHistoryLength) {
      this.operationsHistory.pop();
    }
    
    console.log(`Operation tracked: ${type}`, details);
    console.log("Current operations history:", this.operationsHistory);
  }

  /**
   * Returns the operations history
   * @returns {Array} Array of operations
   */
  getOperationsHistory() {
    // Initialize if undefined
    if (!this.operationsHistory) {
      this.operationsHistory = [];
      console.warn("Operations history was undefined, initialized empty array");
    }
    
    console.log("Getting operations history:", this.operationsHistory);
    return this.operationsHistory;
  }
  
  /**
   * Debug method to directly add a test operation to history
   * This is for testing undo functionality
   */
  addTestOperation() {
    this._trackOperation("test", {
      message: "This is a test operation",
      timestamp: new Date().getTime()
    });
    
    return { success: true, message: "Test operation added to history" };
  }

  /**
   * Clears the operations history
   */
  clearOperationsHistory() {
    this.operationsHistory = [];
  }

  /**
   * Undoes the most recent operation
   * @returns {Promise<Object>} Result of the undo operation
   */
  async undoLastOperation() {
    if (this.operationsHistory.length === 0) {
      return { success: false, message: "No operations to undo" };
    }

    const operation = this.operationsHistory[0];
    console.log("Attempting to undo operation:", operation);

    try {
      let result;
      
      switch (operation.type) {
        case "insertText":
          result = await this._undoInsertText(operation.details);
          break;
        case "insertFormula":
          result = await this._undoInsertFormula(operation.details);
          break;
        case "formatRange":
          result = await this._undoFormatRange(operation.details);
          break;
        case "formatCellsByContent":
          result = await this._undoFormatCellsByContent(operation.details);
          break;
        case "formatRowsByCondition":
          result = await this._undoFormatRowsByCondition(operation.details);
          break;
        case "formatRowsByExactMatch":
          result = await this._undoFormatRowsByExactMatch(operation.details);
          break;
        case "createChart":
          result = await this._undoCreateChart(operation.details);
          break;
        case "createTable":
          result = await this._undoCreateTable(operation.details);
          break;
        case "formatAsTable":
          result = await this._undoFormatAsTable(operation.details);
          break;
        case "createNewWorksheet":
          result = await this._undoCreateNewWorksheet(operation.details);
          break;
        case "renameWorksheet":
          result = await this._undoRenameWorksheet(operation.details);
          break;
        case "addConditionalFormat":
          result = await this._undoAddConditionalFormat(operation.details);
          break;
        case "executeCode":
          result = await this._undoExecuteCode(operation.details);
          break;
        case "test":
          result = await this._undoTestOperation(operation.details);
          break;
        default:
          result = { success: false, message: `Undo not supported for operation type: ${operation.type}` };
      }

      if (result.success) {
        // Remove the operation from history only if undo was successful
        this.operationsHistory.shift();
        // Log the updated history
        console.log("Updated operations history after undo:", this.operationsHistory);
      }

      return {
        success: result.success,
        message: result.success ? `Successfully undid ${operation.type} operation` : result.message,
        details: operation
      };
    } catch (error) {
      console.error("Error undoing operation:", error);
      return {
        success: false,
        message: `Failed to undo operation: ${error.message}`,
        details: operation
      };
    }
  }

  /**
   * Undoes an insertText operation
   * @param {Object} details - Operation details
   * @private
   */
  async _undoInsertText(details) {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(details.address);
        
        // Load current format to preserve it
        range.format.load(["fill", "font"]);
        range.load("values");
        await context.sync();
        
        // Store current formatting
        const currentFormat = {
          fillColor: range.format.fill.color,
          fontColor: range.format.font.color,
          fontBold: range.format.font.bold,
          fontItalic: range.format.font.italic,
          fontUnderline: range.format.font.underline
        };
        
        // If we have previous value, restore it, otherwise clear the cell
        if (details.previousValue !== undefined) {
          range.values = [[details.previousValue]];
        } else {
          range.clear("Contents"); // Only clear contents, not formatting
        }
        
        // Restore the formatting we captured
        range.format.fill.color = currentFormat.fillColor;
        range.format.font.color = currentFormat.fontColor;
        range.format.font.bold = currentFormat.fontBold;
        range.format.font.italic = currentFormat.fontItalic;
        range.format.font.underline = currentFormat.fontUnderline;
        
        await context.sync();
      });
      
      return { success: true };
    } catch (error) {
      console.error("Error undoing insertText:", error);
      return { success: false, message: error.message };
    }
  }

  /**
   * Undoes an insertFormula operation
   * @param {Object} details - Operation details
   * @private
   */
  async _undoInsertFormula(details) {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(details.address);
        
        // Load current format to preserve it
        range.format.load(["fill", "font"]);
        range.load(["values", "formulas"]);
        await context.sync();
        
        // Store current formatting
        const currentFormat = {
          fillColor: range.format.fill.color,
          fontColor: range.format.font.color,
          fontBold: range.format.font.bold,
          fontItalic: range.format.font.italic,
          fontUnderline: range.format.font.underline
        };
        
        // If we have previous value, restore it, otherwise clear the cell
        if (details.previousValue !== undefined) {
          if (details.previousValue && typeof details.previousValue === 'string' && 
              details.previousValue.startsWith && details.previousValue.startsWith("=")) {
            range.formulas = [[details.previousValue]];
          } else {
            range.values = [[details.previousValue]];
          }
        } else {
          range.clear("Contents"); // Only clear contents, not formatting
        }
        
        // Restore the formatting we captured
        range.format.fill.color = currentFormat.fillColor;
        range.format.font.color = currentFormat.fontColor;
        range.format.font.bold = currentFormat.fontBold;
        range.format.font.italic = currentFormat.fontItalic;
        range.format.font.underline = currentFormat.fontUnderline;
        
        await context.sync();
      });
      
      return { success: true };
    } catch (error) {
      console.error("Error undoing insertFormula:", error);
      return { success: false, message: error.message };
    }
  }

  /**
   * Undoes a formatRange operation
   * @param {Object} details - Operation details
   * @private
   */
  async _undoFormatRange(details) {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(details.range);
        
        // Restore previous formatting if available
        if (details.previousFormatting) {
          const format = range.format;
          
          // Restore text formatting
          if (details.previousFormatting.bold !== undefined) format.font.bold = details.previousFormatting.bold;
          if (details.previousFormatting.italic !== undefined) format.font.italic = details.previousFormatting.italic;
          if (details.previousFormatting.underline !== undefined) format.font.underline = details.previousFormatting.underline;
          if (details.previousFormatting.fontSize !== undefined) format.font.size = details.previousFormatting.fontSize;
          
          // Restore colors
          if (details.previousFormatting.fill !== undefined) format.fill.color = details.previousFormatting.fill;
          if (details.previousFormatting.color !== undefined) format.font.color = details.previousFormatting.color;
          
          // Restore borders if they were changed
          if (details.previousFormatting.hasBorders !== undefined) {
            const borderEdges = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"];
            for (const edge of borderEdges) {
              format.borders.getItem(edge).style = details.previousFormatting.hasBorders 
                ? "Continuous" 
                : "None";
            }
          }
          
          // Restore alignment
          if (details.previousFormatting.horizontalAlignment !== undefined) {
            format.horizontalAlignment = details.previousFormatting.horizontalAlignment;
          }
        } else {
          // If no previous formatting recorded, DO NOT clear the formatting
          // as it might remove unrelated formatting
          console.warn("No previous formatting data available for undo");
        }
        
        await context.sync();
      });
      
      return { success: true };
    } catch (error) {
      console.error("Error undoing formatRange:", error);
      return { success: false, message: error.message };
    }
  }

  /**
   * Undoes conditional formatting operations (applies to multiple formats)
   * @param {Object} details - Operation details
   * @private
   */
  async _undoConditionalFormatting(details) {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        
        // If we have affected cells recorded, restore their formatting
        if (details.affectedCells && details.affectedCells.length > 0) {
          console.log(`Restoring formatting for ${details.affectedCells.length} affected cells/rows`);
          for (const cell of details.affectedCells) {
            const range = sheet.getRange(cell.address);
            const format = range.format;
            
            // Reset all the formatting properties we stored
            // Use null to reset to default or apply the previous specific value if available
            if (details.options?.fillColor !== undefined || cell.previousFill !== undefined) {
              format.fill.color = cell.previousFill === undefined ? null : cell.previousFill;
            }
            
            if (details.options?.fontColor !== undefined || cell.previousFontColor !== undefined) {
              format.font.color = cell.previousFontColor === undefined ? null : cell.previousFontColor;
            }
            
            if (details.options?.bold !== undefined || cell.previousBold !== undefined) {
              format.font.bold = cell.previousBold === undefined ? false : cell.previousBold;
            }
            
            if (details.options?.italic !== undefined || cell.previousItalic !== undefined) {
              format.font.italic = cell.previousItalic === undefined ? false : cell.previousItalic;
            }
            
            if (details.options?.underline !== undefined || cell.previousUnderline !== undefined) {
              // Assuming default is 'None' for underline
              format.font.underline = cell.previousUnderline === undefined ? "None" : cell.previousUnderline;
            }
          }
        } else {
          // If no specific cells recorded, do not clear formatting from the entire range
          // as it might remove unrelated formatting
          console.warn("No affected cells data available for undo");
        }
        
        await context.sync();
      });
      
      return { success: true };
    } catch (error) {
      console.error("Error undoing conditional formatting:", error);
      return { success: false, message: error.message };
    }
  }

  /**
   * Undoes a createChart operation
   * @param {Object} details - Operation details
   * @private
   */
  async _undoCreateChart(details) {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        
        // If we have the chart ID, delete it directly
        if (details.chartId) {
          const chart = sheet.charts.getItem(details.chartId);
          chart.delete();
        } else {
          // Otherwise, try to find and delete the chart at the specified position
          const charts = sheet.charts;
          charts.load("items");
          await context.sync();
          
          // Try to identify the chart by position or other criteria
          for (let i = 0; i < charts.items.length; i++) {
            if (charts.items[i].left === details.position.left &&
                charts.items[i].top === details.position.top) {
              charts.items[i].delete();
              break;
            }
          }
        }
        
        await context.sync();
      });
      
      return { success: true };
    } catch (error) {
      console.error("Error undoing createChart:", error);
      return { success: false, message: error.message };
    }
  }

  /**
   * Undoes a createTable operation
   * @param {Object} details - Operation details
   * @private
   */
  async _undoCreateTable(details) {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        
        // If we have the table ID, delete it directly
        if (details.tableId) {
          const table = sheet.tables.getItem(details.tableId);
          table.delete();
        } else if (details.tableName) {
          // Try to find by name
          const table = sheet.tables.getItem(details.tableName);
          table.delete();
        } else if (details.range) {
          // If we have the range, try to get tables in that range
          const range = sheet.getRange(details.range);
          const tables = context.workbook.tables;
          tables.load("items");
          await context.sync();
          
          // Find tables that overlap with our range
          for (let i = 0; i < tables.items.length; i++) {
            const tableRange = tables.items[i].getRange();
            tableRange.load("address");
            await context.sync();
            
            if (tableRange.address === range.address) {
              tables.items[i].delete();
              break;
            }
          }
        }
        
        await context.sync();
      });
      
      return { success: true };
    } catch (error) {
      console.error("Error undoing createTable:", error);
      return { success: false, message: error.message };
    }
  }

  /**
   * Undoes a formatAsTable operation
   * @param {Object} details - Operation details
   * @private
   */
  async _undoFormatAsTable(details) {
    // This is similar to undoing formatRange
    return this._undoFormatRange(details);
  }

  /**
   * Undoes a createNewWorksheet operation
   * @param {Object} details - Operation details
   * @private
   */
  async _undoCreateNewWorksheet(details) {
    try {
      await Excel.run(async (context) => {
        // Get the worksheet by name
        const worksheet = context.workbook.worksheets.getItem(details.name);
        worksheet.delete();
        await context.sync();
      });
      
      return { success: true };
    } catch (error) {
      console.error("Error undoing createNewWorksheet:", error);
      return { success: false, message: error.message };
    }
  }

  /**
   * Undoes a renameWorksheet operation
   * @param {Object} details - Operation details
   * @private
   */
  async _undoRenameWorksheet(details) {
    try {
      await Excel.run(async (context) => {
        // Get the worksheet with the new name and rename it back
        const worksheet = context.workbook.worksheets.getItem(details.newName);
        worksheet.name = details.oldName;
        await context.sync();
      });
      
      return { success: true };
    } catch (error) {
      console.error("Error undoing renameWorksheet:", error);
      return { success: false, message: error.message };
    }
  }

  /**
   * Inserts text into a specified cell
   * @param {string} text - The text to insert
   * @param {string} address - The cell address (e.g., "A1")
   * @returns {Promise} - A promise that resolves when the operation is complete
   */
  async insertText(text, address) {
    try {
      let previousValue;
      
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(address);
        
        // Load the current value for undo purposes
        range.load("values");
        await context.sync();
        
        previousValue = range.values[0][0];
        range.values = [[text]];
        await context.sync();
      });
      
      // Track operation for undo
      this._trackOperation("insertText", {
        address,
        text,
        previousValue
      });
      
      return { success: true };
    } catch (error) {
      console.error("Error inserting text:", error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Inserts a formula into a specified cell
   * @param {string} formula - The formula to insert (without the = sign)
   * @param {string} address - The cell address (e.g., "A1")
   * @returns {Promise} - A promise that resolves when the operation is complete
   */
  async insertFormula(formula, address) {
    try {
      let previousValue;
      
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(address);
        
        // Load the current value for undo purposes
        range.load(["values", "formulas"]);
        await context.sync();
        
        previousValue = range.formulas[0][0];
        
        // Make sure we prefix with = if not already present
        const formulaText = formula.startsWith("=") ? formula : `=${formula}`;
        range.formulas = [[formulaText]];
        await context.sync();
      });
      
      // Track operation for undo
      this._trackOperation("insertFormula", {
        address,
        formula,
        previousValue
      });
      
      return { success: true };
    } catch (error) {
      console.error("Error inserting formula:", error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Retrieves data from a specified range
   * @param {string} range - The range to get data from (e.g., "A1:B10")
   * @returns {Promise<Array>} - A promise that resolves with the retrieved data
   */
  async getData(range) {
    try {
      let result;
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const rangeObj = sheet.getRange(range);
        rangeObj.load("values");
        await context.sync();
        result = rangeObj.values;
      });
      return { success: true, data: result };
    } catch (error) {
      console.error("Error getting data:", error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Gets information about the currently selected range
   * @returns {Promise<Object>} - A promise that resolves with information about the selected range
   */
  async getSelectedRange() {
    try {
      let result = {};
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load(["address", "values", "rowCount", "columnCount", "formulas"]);
        await context.sync();
        
        result = {
          address: range.address,
          values: range.values,
          rowCount: range.rowCount,
          columnCount: range.columnCount,
          formulas: range.formulas
        };
      });
      return { success: true, data: result };
    } catch (error) {
      console.error("Error getting selected range:", error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Creates a chart based on the specified data
   * @param {string} dataRange - The range containing the data for the chart
   * @param {string} chartType - The type of chart to create
   * @returns {Promise} - A promise that resolves when the operation is complete
   */
  async createChart(dataRange, chartType = "ColumnClustered") {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(dataRange);
        
        // Get a range that's 20 rows down to place the chart
        const chartRange = range.getOffsetRange(range.rowCount + 2, 0);
        const chart = sheet.charts.add(
          Excel.ChartType[chartType], 
          range, 
          Excel.ChartSeriesBy.auto
        );
        
        // Set chart properties
        chart.title.text = "Generated Chart";
        chart.setPosition(chartRange.rowIndex, chartRange.columnIndex);
        chart.width = 400;
        chart.height = 300;
        
        await context.sync();
      });
      return { success: true };
    } catch (error) {
      console.error("Error creating chart:", error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Formats a specified range
   * @param {string} range - The range to format
   * @param {Object} formatting - The formatting options to apply
   * @returns {Promise} - A promise that resolves when the operation is complete
   */
  async formatRange(range, formatting) {
    try {
      let previousFormatting = {};
      
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const rangeObj = sheet.getRange(range);
        
        // Load current formatting for undo purposes
        const format = rangeObj.format;
        format.font.load(["bold", "italic", "underline", "size", "color"]);
        format.fill.load("color");
        format.load("horizontalAlignment");
        await context.sync();
        
        // Store current formatting
        previousFormatting = {
          bold: format.font.bold,
          italic: format.font.italic,
          underline: format.font.underline,
          fontSize: format.font.size,
          color: format.font.color,
          fill: format.fill.color,
          horizontalAlignment: format.horizontalAlignment
        };
        
        // Apply text formatting
        if (formatting.bold !== undefined) format.font.bold = formatting.bold;
        if (formatting.italic !== undefined) format.font.italic = formatting.italic;
        if (formatting.underline !== undefined) format.font.underline = formatting.underline;
        if (formatting.fontSize !== undefined) format.font.size = formatting.fontSize;
        
        // Apply colors
        if (formatting.fill !== undefined) {
          // Map common color names to appropriate Excel colors
          const colorMap = {
            red: "#FF0000",
            blue: "#0000FF",
            green: "#00FF00",
            yellow: "#FFFF00",
            orange: "#FFA500",
            purple: "#800080",
            pink: "#FFC0CB",
            brown: "#A52A2A",
            black: "#000000",
            white: "#FFFFFF",
            gray: "#808080",
            grey: "#808080",
            aqua: "#00FFFF",
            cyan: "#00FFFF",
            magenta: "#FF00FF",
            gold: "#FFD700",
            silver: "#C0C0C0",
            violet: "#EE82EE",
            indigo: "#4B0082",
            turquoise: "#40E0D0",
            navy: "#000080",
            teal: "#008080"
          };
          
          format.fill.color = colorMap[formatting.fill.toLowerCase()] || formatting.fill;
        }
        
        if (formatting.color !== undefined) {
          const colorMap = {
            red: "#FF0000",
            blue: "#0000FF",
            green: "#00FF00",
            yellow: "#FFFF00",
            orange: "#FFA500",
            purple: "#800080",
            pink: "#FFC0CB",
            brown: "#A52A2A",
            black: "#000000",
            white: "#FFFFFF",
            gray: "#808080",
            grey: "#808080",
            aqua: "#00FFFF",
            cyan: "#00FFFF",
            magenta: "#FF00FF",
            gold: "#FFD700",
            silver: "#C0C0C0",
            violet: "#EE82EE",
            indigo: "#4B0082",
            turquoise: "#40E0D0",
            navy: "#000080",
            teal: "#008080"
          };
          
          format.font.color = colorMap[formatting.color.toLowerCase()] || formatting.color;
        }
        
        // Apply borders
        if (formatting.border) {
          // Track border state for undo
          previousFormatting.hasBorders = format.borders.getItem("EdgeTop").style !== "None";
          
          format.borders.getItem("EdgeTop").style = "Continuous";
          format.borders.getItem("EdgeBottom").style = "Continuous";
          format.borders.getItem("EdgeLeft").style = "Continuous";
          format.borders.getItem("EdgeRight").style = "Continuous";
        }
        
        // Apply alignment
        if (formatting.horizontalAlignment) {
          format.horizontalAlignment = formatting.horizontalAlignment;
        }
        
        await context.sync();
      });
      
      // Track operation for undo
      this._trackOperation("formatRange", {
        range,
        formatting,
        previousFormatting
      });
      
      return { success: true };
    } catch (error) {
      console.error("Error formatting range:", error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Creates a new worksheet
   * @param {string} name - Optional name for the new worksheet
   * @returns {Promise} - A promise that resolves when the operation is complete
   */
  async createNewWorksheet(name = null) {
    try {
      let worksheetName;
      
      await Excel.run(async (context) => {
        const newSheet = context.workbook.worksheets.add();
        
        if (name) {
          newSheet.name = name;
          worksheetName = name;
        } else {
          // Load the auto-generated name
          newSheet.load("name");
        }
        
        newSheet.activate();
        await context.sync();
        
        if (!name) {
          // Capture the auto-generated name
          worksheetName = newSheet.name;
        }
      });
      
      // Track operation for undo
      this._trackOperation("createNewWorksheet", {
        name: worksheetName
      });
      
      return { success: true };
    } catch (error) {
      console.error("Error creating new worksheet:", error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Renames a worksheet
   * @param {string} oldName - The current name of the worksheet
   * @param {string} newName - The new name for the worksheet
   * @returns {Promise} - A promise that resolves when the operation is complete
   */
  async renameWorksheet(oldName, newName) {
    try {
      await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getItem(oldName);
        worksheet.name = newName;
        await context.sync();
      });
      
      // Track operation for undo
      this._trackOperation("renameWorksheet", {
        oldName,
        newName
      });
      
      return { success: true };
    } catch (error) {
      console.error("Error renaming worksheet:", error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Calculates the sum of a range
   * @param {string} range - The range to calculate the sum of
   * @returns {Promise<number>} - A promise that resolves with the sum
   */
  async calculateSum(range) {
    try {
      let result;
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const rangeObj = sheet.getRange(range);
        rangeObj.load("values");
        await context.sync();
        
        // Calculate the sum
        result = rangeObj.values.reduce((sum, row) => {
          return sum + row.reduce((rowSum, cell) => {
            const value = parseFloat(cell);
            return rowSum + (isNaN(value) ? 0 : value);
          }, 0);
        }, 0);
        
        // Track how many cells were numeric vs non-numeric
        let numericCount = 0;
        let nonNumericCount = 0;
        rangeObj.values.flat().forEach(cell => {
          if (cell !== null && cell !== "") {
            const value = parseFloat(cell);
            if (!isNaN(value)) {
              numericCount++;
            } else {
              nonNumericCount++;
            }
          }
        });
        
        calculationResult = {
          sum: result,
          numericCells: numericCount,
          nonNumericCells: nonNumericCount,
          cellCount: rangeObj.values.flat().length,
          address: rangeObj.address
        };
        
      });
      // Don't track simple calculations in history for now
      // this._trackOperation("calculateSum", { range, result: calculationResult });
      return { success: true, result: calculationResult };
    } catch (error) {
      console.error("Error calculating sum:", error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Retrieves all data from the active worksheet
   * @returns {Promise<Object>} - A promise that resolves with all worksheet data
   */
  async getAllData() {
    try {
      let result = {};
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRange();
        usedRange.load(["address", "values", "rowCount", "columnCount"]);
        await context.sync();
        
        // Check if the worksheet is empty
        const isEmpty = usedRange.rowCount === 0 || usedRange.columnCount === 0;
        
        result = {
          address: usedRange.address,
          values: usedRange.values,
          rowCount: usedRange.rowCount,
          columnCount: usedRange.columnCount,
          isEmpty: isEmpty
        };
      });
      return { success: true, ...result };
    } catch (error) {
      console.error("Error getting all worksheet data:", error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Creates a formatted table with headers from the provided data
   * @param {Array<Array<any>>} data - 2D array of data where the first row is treated as headers
   * @param {string} startCell - The starting cell for the table (default: "A1")
   * @param {Object} options - Additional options for the table
   * @param {string} options.tableName - Optional name for the table
   * @param {string} options.tableStyle - Table style (default: "TableStyleMedium2")
   * @param {boolean} options.hasTotals - Whether to add a totals row (default: false)
   * @returns {Promise<Object>} - A promise that resolves with the operation result
   */
  async createTable(data, startCell = "A1", options = {}) {
    try {
      if (!data || !Array.isArray(data) || data.length === 0) {
        throw new Error("Invalid data: expected non-empty 2D array");
      }

      const tableName = options.tableName || `Table${Date.now()}`;
      const tableStyle = options.tableStyle || "TableStyleMedium2";
      const hasTotals = options.hasTotals || false;

      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        
        // Calculate the table range
        const rowCount = data.length;
        const columnCount = Math.max(...data.map(row => row.length));
        
        // Parse the starting cell to get row and column indices
        const columnMatch = startCell.match(/[A-Za-z]+/)[0];
        const rowMatch = parseInt(startCell.match(/\d+/)[0]);
        
        // Convert column letters to column index (A=1, B=2, etc.)
        let startColumnIndex = 0;
        for (let i = 0; i < columnMatch.length; i++) {
          startColumnIndex = startColumnIndex * 26 + columnMatch.charCodeAt(i) - 64;
        }
        
        // Calculate the end cell
        const endColumnIndex = startColumnIndex + columnCount - 1;
        const endRowIndex = rowMatch + rowCount - 1;
        
        // Convert the end column index back to a letter
        let endColumnLetter = "";
        let tempColumnIndex = endColumnIndex;
        while (tempColumnIndex > 0) {
          const remainder = (tempColumnIndex - 1) % 26;
          endColumnLetter = String.fromCharCode(65 + remainder) + endColumnLetter;
          tempColumnIndex = Math.floor((tempColumnIndex - 1) / 26);
        }
        
        const tableRange = `${columnMatch}${rowMatch}:${endColumnLetter}${endRowIndex}`;
        
        // Get the range and set the values
        const range = sheet.getRange(tableRange);
        range.values = data;
        
        // Create a table from the range
        const table = sheet.tables.add(tableRange, true);
        table.name = tableName;
        table.style = tableStyle;
        table.showTotals = hasTotals;
        
        // Format headers with bold
        const headerRange = sheet.getRange(`${columnMatch}${rowMatch}:${endColumnLetter}${rowMatch}`);
        headerRange.format.font.bold = true;
        
        // Auto-fit columns
        range.format.autofitColumns();
        
        await context.sync();
      });
      
      return { success: true };
    } catch (error) {
      console.error("Error creating table:", error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Formats a specified range as a professional looking table
   * @param {string} range - The range to format (e.g., "A1:L11")
   * @param {Object} options - Formatting options
   * @param {boolean} options.hasHeaders - Whether the first row contains headers (default: true)
   * @param {string} options.headerFill - Background color for headers (default: "navy")
   * @param {string} options.headerFont - Font color for headers (default: "white")
   * @param {string} options.alternateFill - Background color for alternate rows (default: null)
   * @param {boolean} options.autofitColumns - Whether to autofit columns (default: true)
   * @returns {Promise} - A promise that resolves when the operation is complete
   */
  async formatAsTable(range, options = {}) {
    try {
      const hasHeaders = options.hasHeaders !== undefined ? options.hasHeaders : true;
      const headerFill = options.headerFill || "navy";
      const headerFont = options.headerFont || "white";
      const alternateFill = options.alternateFill || null;
      const autofitColumns = options.autofitColumns !== undefined ? options.autofitColumns : true;
      
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const rangeObj = sheet.getRange(range);
        
        // Get information about the range
        rangeObj.load(["rowCount", "columnCount", "address"]);
        await context.sync();
        
        const rowCount = rangeObj.rowCount;
        const columnCount = rangeObj.columnCount;
        
        // Parse the range to get start and end cells
        const rangeMatch = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
        if (!rangeMatch) throw new Error(`Invalid range format: ${range}`);
        
        const startColumn = rangeMatch[1];
        const startRow = parseInt(rangeMatch[2]);
        const endColumn = rangeMatch[3];
        const endRow = parseInt(rangeMatch[4]);
        
        // Format headers if applicable
        if (hasHeaders) {
          const headerRange = sheet.getRange(`${startColumn}${startRow}:${endColumn}${startRow}`);
          
          // Apply header styling
          headerRange.format.fill.color = headerFill;
          headerRange.format.font.color = headerFont;
          headerRange.format.font.bold = true;
          headerRange.format.horizontalAlignment = "Center";
          headerRange.format.verticalAlignment = "Center";
        }
        
        // Apply borders to the entire range
        rangeObj.format.borders.getItem('EdgeTop').style = 'Continuous';
        rangeObj.format.borders.getItem('EdgeBottom').style = 'Continuous';
        rangeObj.format.borders.getItem('EdgeLeft').style = 'Continuous';
        rangeObj.format.borders.getItem('EdgeRight').style = 'Continuous';
        rangeObj.format.borders.getItem('InsideHorizontal').style = 'Continuous';
        rangeObj.format.borders.getItem('InsideVertical').style = 'Continuous';
        
        // Apply alternate row coloring if specified
        if (alternateFill) {
          const startAltRow = hasHeaders ? startRow + 1 : startRow;
          for (let i = startAltRow; i <= endRow; i += 2) {
            const altRowRange = sheet.getRange(`${startColumn}${i}:${endColumn}${i}`);
            altRowRange.format.fill.color = alternateFill;
          }
        }
        
        // Center align all data
        rangeObj.format.horizontalAlignment = "Center";
        rangeObj.format.verticalAlignment = "Center";
        
        // Auto-fit columns if requested
        if (autofitColumns) {
          rangeObj.format.autofitColumns();
        }
        
        await context.sync();
      });
      
      return { success: true };
    } catch (error) {
      console.error("Error formatting as table:", error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Gets the address of the used range in the active worksheet
   * @returns {Promise<Object>} - A promise that resolves with the used range address or an error
   */
  async getUsedRange() {
    try {
      let result = {};
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRange();
        usedRange.load("address");
        await context.sync();
        
        // Check if the worksheet is effectively empty
        if (!usedRange.address || !usedRange.address.includes('!')) { // Check if address is valid
          result = { address: null, isEmpty: true };
        } else {
          result = { address: usedRange.address, isEmpty: false };
        }
      });
      return { success: true, ...result };
    } catch (error) {
      // Handle cases where getUsedRange might fail on a truly empty sheet
      if (error.code === 'ItemNotFound' || error.code === 'GeneralException') { // Added GeneralException check
        return { success: true, address: null, isEmpty: true };
      }
      console.error("Error getting used range:", error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Executes arbitrary Office.js JavaScript code in the Excel context
   * @param {string} codeString - The Office.js code to execute (should include its own Excel.run)
   * @returns {Promise<Object>} - A promise that resolves with the execution result
   */
  async executeOfficeJsCode(codeString) {
    try {
      console.log("Attempting to execute Office.js code string:", codeString);

      // Check if Excel is defined globally
      if (typeof Excel === 'undefined') {
        console.error("Excel is not defined globally. Office.js may not be fully loaded.");
        return { success: false, error: "Excel is not defined globally. Office.js may not be fully loaded." };
      }

      // Create an async function from the code string.
      // Pass necessary globals (Excel, console) into the function's scope.
      // The codeString itself is expected to contain the await Excel.run().
      const asyncFunction = new Function('Excel', 'console', `return (async () => { ${codeString} })();`);

      // Execute the function, passing the global Excel and console objects.
      await asyncFunction(Excel, console);

      // Track this operation in history
      this._trackOperation("executeCode", {
        code: codeString,
        timestamp: new Date().getTime()
      });

      console.log("Office.js code string execution completed successfully.");
      return { success: true };
    } catch (error) {
      console.error("Error executing Office.js code string:", error);
      let errorDetails = {
        message: error.message,
        name: error.name,
        stack: error.stack,
      };
      console.log("Error details:", errorDetails);
      return { success: false, error: error.message, details: errorDetails };
    }
  }

  /**
   * Adds conditional formatting to a range based on specific criteria
   * @param {string} range - The range to apply conditional formatting to
   * @param {object} criteria - The criteria for the conditional formatting
   * @param {string} criteria.type - The type of conditional format ("containsText", "cellValue", "colorScale", etc.)
   * @param {string} criteria.operator - The operator for the condition (e.g., "contains", "equals", "greaterThan")
   * @param {string|number} criteria.value - The value to compare against
   * @param {object} format - The formatting to apply
   * @returns {Promise<object>} - Result of the operation
   */
  async addConditionalFormat(range, criteria, format) {
    try {
      let formatId;
      
      await Excel.run(function(context) { // Use function(context)
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const rangeObj = sheet.getRange(range);
        
        // Use conditionalFormats (plural) instead of conditionalFormat (singular)
        const conditionalFormats = rangeObj.conditionalFormats;
        let conditionalFormat;
        
        // Apply the appropriate conditional format based on the criteria type
        switch (criteria.type) {
          case "containsText":
            conditionalFormat = conditionalFormats.add(Excel.ConditionalFormatType.containsText); // Correct Enum
            conditionalFormat.textComparison.rule = { operator: criteria.operator || "Contains", text: criteria.value }; // Allow operator override
            break;
            
          case "cellValue":
            conditionalFormat = conditionalFormats.add(Excel.ConditionalFormatType.cellValue); // Correct Enum
            conditionalFormat.cellValue.rule = {
              formula1: criteria.value,
              operator: criteria.operator || Excel.ConditionalCellValueOperator.equalTo // Use Excel Enum for operator
            };
            break;
            
          case "colorScale":
            conditionalFormat = conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
            // Configure the color scale...
            break;
            
          case "dataBar":
            conditionalFormat = conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
            // Configure the data bar...
            break;
            
          case "iconSet":
            conditionalFormat = conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
            // Configure the icon set...
            break;
            
          case "topBottom":
            conditionalFormat = conditionalFormats.add(Excel.ConditionalFormatType.topBottom);
            conditionalFormat.topBottom.rule = {
              rank: criteria.value || 10, // Ensure value exists
              type: criteria.operator || Excel.ConditionalTopBottomCriterionType.topItems // Correct Enum
            };
            break;
            
          default:
            // Use console.error and reject the promise for unsupported types
            console.error(`Unsupported conditional format type: ${criteria.type}`);
            // Reject the promise within Excel.run context if possible, or handle outside
            // For simplicity, we'll let it potentially fail at sync or handle after run
            throw new Error(`Unsupported conditional format type: ${criteria.type}`);
        }
        
        // Apply the formatting if format object and conditionalFormat exist
        if (format && conditionalFormat) {
          if (format.fillColor) conditionalFormat.format.fill.color = format.fillColor;
          if (format.fontColor) conditionalFormat.format.font.color = format.fontColor;
          if (format.bold !== undefined) conditionalFormat.format.font.bold = format.bold;
          // Add italic, underline etc. if needed
        } else if (!conditionalFormat) {
          throw new Error(`Conditional format could not be created for type: ${criteria.type}`);
        }
        
        // Load the format ID for tracking
        conditionalFormat.load("id"); // Load ID for tracking
        
        return context.sync() // Use return context.sync()
          .then(function() {
            formatId = conditionalFormat.id; // Get ID after sync
            console.log(`Conditional format added to range ${range}. ID: ${formatId}`);
          });
          // Catch is handled by the outer try/catch for the async function
      });
      
      // Track operation *after* Excel.run succeeds
      // Track this operation for undo
      this._trackOperation("addConditionalFormat", {
        range,
        criteria,
        format,
        formatId // Pass the retrieved ID
      });
      
      return { 
        success: true,
        message: `Successfully applied conditional formatting to ${range}`
      };
    } catch (error) {
      console.error("Error adding conditional format:", error);
      // Check if it's the specific unsupported type error
      if (error.message.startsWith("Unsupported conditional format type:")) {
        return { success: false, error: error.message };
      }
      // Provide more specific error details if available from Office.js
      const debugInfo = error.debugInfo ? JSON.stringify(error.debugInfo) : "";
      return { success: false, error: `Failed to add conditional format: ${error.message} ${debugInfo}`.trim() };
    }
  }

  /**
   * Undoes an addConditionalFormat operation
   * @param {Object} details - Operation details
   * @private
   */
  async _undoAddConditionalFormat(details) {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const rangeObj = sheet.getRange(details.range);
        
        // If we have the format ID, remove that specific format
        if (details.formatId) {
          // Note: using conditionalFormats (plural)
          const conditionalFormats = rangeObj.conditionalFormats;
          conditionalFormats.getItem(details.formatId).delete();
        } else {
          // If no specific ID, clear all conditional formats from the range
          rangeObj.conditionalFormats.clearAll();
        }
        
        await context.sync();
      });
      
      return { success: true };
    } catch (error) {
      console.error("Error undoing conditional format:", error);
      return { success: false, message: error.message };
    }
  }

  /**
   * Undoes an executeCode operation
   * @param {Object} details - Operation details
   * @private
   */
  async _undoExecuteCode(details) {
    console.warn("Attempting to undo 'executeCode'. This is generally not supported.", details);
    // Reliable undo for arbitrary code execution is not feasible without
    // knowing the exact inverse operation or storing extensive state,
    // which isn't done for generic code blocks.
    return {
      success: false, // Indicate that the state wasn't actually reverted
      message: "Undo is not supported for actions generated automatically by the AI assistant. Please manually revert the changes if necessary."
    };
  }

  /**
   * Undoes a test operation
   * @param {Object} details - Operation details
   * @private
   */
  async _undoTestOperation(details) {
    try {
      console.log("Undoing test operation:", details);
      // Since this is just a test operation, we don't need to do anything except return success
      return { success: true, message: "Test operation undone successfully" };
    } catch (error) {
      console.error("Error undoing test operation:", error);
      return { success: false, message: error.message };
    }
  }

  /**
   * Undoes conditional formatting operations (applies to multiple formats)
   * @param {Object} details - Operation details
   * @private
   */
  async _undoFormatCellsByContent(details) {
    return this._undoConditionalFormatting(details);
  }

  async _undoFormatRowsByCondition(details) {
    return this._undoConditionalFormatting(details);
  }

  async _undoFormatRowsByExactMatch(details) {
    return this._undoConditionalFormatting(details);
  }

  /**
   * Sorts a specified range based on one or more columns
   * @param {string} rangeAddress - The range address to sort (e.g., "A1:D10")
   * @param {Array<object>} sortFields - Array of sort fields, e.g., [{ key: 0, ascending: true }, { key: 1, ascending: false }] (key is 0-based column index within the range)
   * @param {boolean} hasHeaders - Does the range include headers?
   * @returns {Promise<Object>} - Result of the sort operation
   */
  async sortRange(rangeAddress, sortFields, hasHeaders = true) {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(rangeAddress);
        
        // Apply sort
        range.sort.apply(sortFields, hasHeaders);
        
        await context.sync();
      });
      
      // Track operation for undo (Note: Undoing sort is complex, might just log it)
      this._trackOperation("sortRange", {
        range: rangeAddress,
        sortFields,
        hasHeaders
      });
      
      return { success: true };
    } catch (error) {
      console.error("Error sorting range:", error);
      return { success: false, error: error.message };
    }
  }
  
  /**
   * Applies an AutoFilter to a specified range or the entire worksheet
   * @param {string|null} rangeAddress - The range address to filter (e.g., "A1:D10") or null to apply to the used range.
   * @param {number} columnIndex - The 0-based column index within the range to apply the filter criterion.
   * @param {object} criteria - The filter criteria (e.g., { filterOn: "Values", values: ["Value1", "Value2"] }).
   * See Office.js documentation for Excel.FilterCriteria structure.
   * @returns {Promise<Object>} - Result of the filter operation
   */
  async applyFilter(rangeAddress, columnIndex, criteria) {
    try {
      let appliedRangeAddress;
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        let range;
        
        if (rangeAddress) {
          range = sheet.getRange(rangeAddress);
          appliedRangeAddress = rangeAddress;
        } else {
          // If no range is provided, apply to the used range of the sheet
          range = sheet.getUsedRange();
          range.load("address");
          await context.sync();
          appliedRangeAddress = range.address;
        }
        
        // Check if a filter already exists and clear it before applying a new one
        // Note: This simple check might not cover all cases. A more robust check
        // might involve loading filter properties.
        if (sheet.autoFilter && sheet.autoFilter.enabled) {
          sheet.autoFilter.remove();
          await context.sync(); // Sync after removal
        }
        
        // Apply the filter
        sheet.autoFilter.apply(range, columnIndex, criteria);
        
        await context.sync();
      });
      
      // Track operation for undo (Undoing filters involves removing them)
      this._trackOperation("applyFilter", {
        range: appliedRangeAddress,
        columnIndex,
        criteria
      });
      
      return { success: true };
    } catch (error) {
      console.error("Error applying filter:", error);
      return { success: false, error: error.message };
    }
  }
  
  /**
   * Removes the AutoFilter from the active worksheet
   * @returns {Promise<Object>} - Result of the operation
   */
  async removeFilter() {
    try {
      let filterRange = null;
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        
        // Check if autofilter exists before trying to remove
        sheet.autoFilter.load('enabled, range/address');
        await context.sync();
        
        if (sheet.autoFilter.enabled) {
          filterRange = sheet.autoFilter.range.address; // Store range for potential undo
          sheet.autoFilter.remove();
          await context.sync();
        } else {
          console.log("No active filter to remove.");
        }
      });
      
      // Track operation for undo (Undoing remove means reapplying the filter, which is hard)
      // For simplicity, we might just log or prevent undo for this.
      // If we stored the criteria in 'applyFilter', we could potentially reapply.
      this._trackOperation("removeFilter", {
        previousFilterRange: filterRange
      });
      
      return { success: true };
    } catch (error) {
      console.error("Error removing filter:", error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Gets context about the current state of the workbook and active sheet.
   * @returns {Promise<Object>} - Object containing context information or error.
   */
  async getWorksheetContext() {
    try {
      let contextInfo = {};
      await Excel.run(async (context) => {
        const workbook = context.workbook;
        const worksheets = workbook.worksheets;
        worksheets.load("items/name");
        
        const activeSheet = worksheets.getActiveWorksheet();
        activeSheet.load("name, position, visibility");
        
        const usedRange = activeSheet.getUsedRange(true); // Use true for valuesOnly param
        usedRange.load("address, rowCount, columnCount");
        
        // Try loading sample data (first few rows/cols) - might fail on empty sheets
        let sampleData = null;
        try {
          const sampleRange = activeSheet.getRangeByIndexes(0, 0, Math.min(5, usedRange.rowCount), Math.min(5, usedRange.columnCount));
          sampleRange.load("values");
          await context.sync(); // Sync early to get sample data or handle error
          sampleData = sampleRange.values;
        } catch (sampleError) {
          console.log("Could not load sample data (sheet might be empty or too small).", sampleError);
        }
        
        const selection = workbook.getSelectedRange();
        selection.load("address, values, rowCount, columnCount");
        
        const tables = activeSheet.tables;
        tables.load("items/name");
        
        const charts = activeSheet.charts;
        charts.load("items/name");
        
        await context.sync(); // Final sync for all loaded properties
        
        contextInfo = {
          workbook: {
            name: workbook.name, // Name might not be available depending on context
          },
          worksheets: {
            count: worksheets.items.length,
            names: worksheets.items.map(sheet => sheet.name),
          },
          activeSheet: {
            name: activeSheet.name,
            position: activeSheet.position,
            visibility: activeSheet.visibility,
            usedRange: {
              address: usedRange.address,
              rowCount: usedRange.rowCount,
              columnCount: usedRange.columnCount,
            },
            sampleData: sampleData, // Include sample data
          },
          selection: {
            address: selection.address,
            values: selection.values, // Include values if available
            rowCount: selection.rowCount,
            columnCount: selection.columnCount,
          },
          tables: {
            count: tables.items.length,
            names: tables.items.map(table => table.name),
          },
          charts: {
            count: charts.items.length,
            names: charts.items.map(chart => chart.name),
          },
        };
      });
      return { success: true, context: contextInfo };
    } catch (error) {
      console.error("Error getting worksheet context:", error);
      // Handle specific error for empty sheet used range
      if (error.code === 'ItemNotFound' || error.message.includes("Used range is not available")) {
        // Try to get basic info even if usedRange fails
        try {
          let basicContext = {};
           await Excel.run(async (context) => {
             const workbook = context.workbook;
             const activeSheet = workbook.worksheets.getActiveWorksheet();
             activeSheet.load("name");
             await context.sync();
             basicContext = {
               activeSheet: { name: activeSheet.name, usedRange: { address: "Empty", rowCount: 0, columnCount: 0 } },
               // Add other default/empty values
             };
           });
           return { success: true, context: { ...basicContext, isEmpty: true } };
        } catch (fallbackError) {
          return { success: false, error: fallbackError.message };
        }
      }
      return { success: false, error: error.message };
    }
  }

  /**
   * Retrieves a list of all worksheet names in the workbook.
   * @returns {Promise<Object>} - Object with success status and array of names or error.
   */
  async getWorksheetNames() {
    try {
      let names = [];
      await Excel.run(async (context) => {
        const worksheets = context.workbook.worksheets;
        worksheets.load("items/name");
        await context.sync();
        names = worksheets.items.map(sheet => sheet.name);
      });
      return { success: true, names: names };
    } catch (error) {
      console.error("Error getting worksheet names:", error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Calculates the average of the numeric values in the selected range.
   * @returns {Promise<Object>} - Result object with average, counts, and address.
   */
  async calculateAverage() {
    try {
      let calculationResult = {};
      await Excel.run(async (context) => {
        const rangeObj = context.workbook.getSelectedRange();
        rangeObj.load("values, address");
        await context.sync();
        
        let sum = 0;
        let numericCount = 0;
        let nonNumericCount = 0;
        const values = rangeObj.values.flat();
        
        values.forEach(cell => {
          if (cell !== null && cell !== "") {
            const value = parseFloat(cell);
            if (!isNaN(value)) {
              sum += value;
              numericCount++;
            } else {
              nonNumericCount++;
            }
          }
        });
        
        const average = numericCount > 0 ? sum / numericCount : 0;
        
        calculationResult = {
          average: average,
          numericCells: numericCount,
          nonNumericCells: nonNumericCount,
          cellCount: values.length,
          address: rangeObj.address
        };
      });
      // Don't track for undo
      return { success: true, result: calculationResult };
    } catch (error) {
      console.error("Error calculating average:", error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Counts different types of cells in the selected range.
   * @returns {Promise<Object>} - Result object with counts and address.
   */
  async countCells() {
    try {
      let calculationResult = {};
      await Excel.run(async (context) => {
        const rangeObj = context.workbook.getSelectedRange();
        rangeObj.load("values, address, cellCount"); // Load cellCount for total
        await context.sync();
        
        let numberCells = 0;
        let textCells = 0;
        let blankCells = 0;
        let nonEmptyCells = 0;
        const values = rangeObj.values.flat();
        
        values.forEach(cell => {
          if (cell === null || cell === "") {
            blankCells++;
          } else {
            nonEmptyCells++;
            const value = parseFloat(cell);
            if (!isNaN(value)) {
              numberCells++;
            } else if (typeof cell === 'string') {
              textCells++;
            }
          }
        });
        
        calculationResult = {
          totalCells: rangeObj.cellCount,
          numberCells: numberCells,
          textCells: textCells,
          blankCells: blankCells,
          nonEmptyCells: nonEmptyCells,
          address: rangeObj.address
        };
      });
      // Don't track for undo
      return { success: true, result: calculationResult };
    } catch (error) {
      console.error("Error counting cells:", error);
      return { success: false, error: error.message };
    }
  }
  
  /**
   * Calculates various statistics (sum, average, median, min, max, count) for the selected range.
   * @returns {Promise<Object>} - Result object with statistics and address.
   */
  async getStatistics() {
    try {
      let statsResult = {};
      await Excel.run(async (context) => {
        const rangeObj = context.workbook.getSelectedRange();
        rangeObj.load("values, address");
        await context.sync();

        const numericValues = rangeObj.values
          .flat()
          .map(cell => parseFloat(cell))
          .filter(value => !isNaN(value));

        let sum = 0;
        let min = null;
        let max = null;
        let count = numericValues.length;

        if (count > 0) {
          numericValues.sort((a, b) => a - b);
          sum = numericValues.reduce((acc, val) => acc + val, 0);
          min = numericValues[0];
          max = numericValues[count - 1];
          
          // Calculate median
          const mid = Math.floor(count / 2);
          const median = count % 2 !== 0 ? numericValues[mid] : (numericValues[mid - 1] + numericValues[mid]) / 2;

          statsResult = {
            sum: sum,
            average: sum / count,
            median: median,
            min: min,
            max: max,
            count: count,
            address: rangeObj.address
          };
        } else {
           statsResult = {
            sum: 0, average: 0, median: null, min: null, max: null, count: 0, address: rangeObj.address
           };
        }
      });
      // Don't track for undo
      return { success: true, result: statsResult };
    } catch (error) {
      console.error("Error getting statistics:", error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Undoes a sortRange operation
   * @param {Object} details - Operation details
   * @private
   */
  async _undoSortRange(details) {
    // Undoing a sort requires storing the original order, which can be memory-intensive.
    // For now, we'll log a message indicating that undo is not fully supported.
    console.warn("Undo for sortRange is not fully implemented. Original data order not restored.", details);
    return { success: true, message: "Undo for sort operation is not fully supported." };
  }

  /**
   * Undoes an applyFilter operation by removing the filter
   * @param {Object} details - Operation details
   * @private
   */
  async _undoApplyFilter(details) {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        // If an autofilter exists, remove it
        sheet.autoFilter.load('enabled');
        await context.sync();
        if (sheet.autoFilter.enabled) {
          sheet.autoFilter.remove();
          await context.sync();
        }
      });
      return { success: true };
    } catch (error) {
      console.error("Error undoing applyFilter (removing filter):", error);
      return { success: false, message: error.message };
    }
  }
  
  /**
   * Undoes a removeFilter operation
   * @param {Object} details - Operation details
   * @private
   */
  async _undoRemoveFilter(details) {
     // Re-applying the exact previous filter state is complex.
     // For now, indicate it's not fully supported.
     console.warn("Undo for removeFilter is not fully implemented. Filter not reapplied.", details);
     return { success: true, message: "Undo for filter removal is not fully supported." };
  }

  /**
   * Inserts one or more rows above the specified row index.
   * @param {number} rowIndex - The 0-based index of the row *before* which to insert.
   * @param {number} count - Number of rows to insert (default: 1).
   * @returns {Promise<Object>}
   */
  async insertRows(rowIndex, count = 1) {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        // Get the entire row to insert before
        const referenceRow = sheet.getRangeByIndexes(rowIndex, 0, 1, 0).getEntireRow();
        referenceRow.insert(Excel.InsertShiftDirection.down);
        
        // If inserting multiple rows, repeat the insertion
        // Note: There might be more efficient ways for bulk insertion if needed.
        if (count > 1) {
          for (let i = 1; i < count; i++) {
            // Re-get the reference row as indices shift after insertion
            const nextReferenceRow = sheet.getRangeByIndexes(rowIndex + i, 0, 1, 0).getEntireRow();
            nextReferenceRow.insert(Excel.InsertShiftDirection.down);
          }
        }
        await context.sync();
      });
      
      // Track operation (Undo means deleting the inserted rows)
      this._trackOperation("insertRows", {
        rowIndex: rowIndex,
        count: count
      });
      
      return { success: true };
    } catch (error) {
      console.error("Error inserting rows:", error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Deletes one or more rows starting at the specified row index.
   * @param {number} rowIndex - The 0-based index of the first row to delete.
   * @param {number} count - Number of rows to delete (default: 1).
   * @returns {Promise<Object>}
   */
  async deleteRows(rowIndex, count = 1) {
    try {
       // Note: Storing deleted data for undo is complex.
       // We might skip storing data for simplicity.
      let deletedData = null; // Placeholder for potential future data storage
      
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const rangeToDelete = sheet.getRangeByIndexes(rowIndex, 0, count, 0).getEntireRow();
        
        // // Optional: Load data before deleting for potential undo
        // rangeToDelete.load("values");
        // await context.sync();
        // deletedData = rangeToDelete.values;
        
        rangeToDelete.delete(Excel.DeleteShiftDirection.up);
        await context.sync();
      });
      
      // Track operation (Undo means re-inserting rows, possibly with data)
      this._trackOperation("deleteRows", {
        rowIndex: rowIndex,
        count: count,
        // deletedData: deletedData // Store if implementing data restoration
      });
      
      return { success: true };
    } catch (error) {
      console.error("Error deleting rows:", error);
      return { success: false, error: error.message };
    }
  }
  
  /**
   * Inserts one or more columns to the left of the specified column index.
   * @param {number} colIndex - The 0-based index of the column *before* which to insert.
   * @param {number} count - Number of columns to insert (default: 1).
   * @returns {Promise<Object>}
   */
  async insertColumns(colIndex, count = 1) {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const referenceCol = sheet.getRangeByIndexes(0, colIndex, 0, 1).getEntireColumn();
        referenceCol.insert(Excel.InsertShiftDirection.right);
        
        if (count > 1) {
          for (let i = 1; i < count; i++) {
             const nextReferenceCol = sheet.getRangeByIndexes(0, colIndex + i, 0, 1).getEntireColumn();
             nextReferenceCol.insert(Excel.InsertShiftDirection.right);
          }
        }
        await context.sync();
      });
      
      this._trackOperation("insertColumns", {
        colIndex: colIndex,
        count: count
      });
      
      return { success: true };
    } catch (error) {
      console.error("Error inserting columns:", error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Deletes one or more columns starting at the specified column index.
   * @param {number} colIndex - The 0-based index of the first column to delete.
   * @param {number} count - Number of columns to delete (default: 1).
   * @returns {Promise<Object>}
   */
  async deleteColumns(colIndex, count = 1) {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const rangeToDelete = sheet.getRangeByIndexes(0, colIndex, 0, count).getEntireColumn();
        rangeToDelete.delete(Excel.DeleteShiftDirection.left);
        await context.sync();
      });
      
      this._trackOperation("deleteColumns", {
        colIndex: colIndex,
        count: count
      });
      
      return { success: true };
    } catch (error) {
      console.error("Error deleting columns:", error);
      return { success: false, error: error.message };
    }
  }
  
  /**
   * Hides one or more rows.
   * @param {number} rowIndex - The 0-based index of the first row to hide.
   * @param {number} count - Number of rows to hide (default: 1).
   * @returns {Promise<Object>}
   */
  async hideRows(rowIndex, count = 1) {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const rangeToHide = sheet.getRangeByIndexes(rowIndex, 0, count, 0).getEntireRow();
        rangeToHide.format.rowHeight = 0;
        await context.sync();
      });
      
      this._trackOperation("hideRows", {
        rowIndex: rowIndex,
        count: count
      });
      
      return { success: true };
    } catch (error) {
      console.error("Error hiding rows:", error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Unhides one or more rows.
   * @param {number} rowIndex - The 0-based index of the first row to unhide.
   * @param {number} count - Number of rows to unhide (default: 1).
   * @returns {Promise<Object>}
   */
  async unhideRows(rowIndex, count = 1) {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const rangeToUnhide = sheet.getRangeByIndexes(rowIndex, 0, count, 0).getEntireRow();
        // Setting rowHeight to null or undefined might reset to default, 
        // but explicitly setting visible is clearer.
        rangeToUnhide.rowHidden = false;
        await context.sync();
      });
      
      // Track operation (Undo means hiding again)
      this._trackOperation("unhideRows", {
        rowIndex: rowIndex,
        count: count
      });
      
      return { success: true };
    } catch (error) {
      console.error("Error unhiding rows:", error);
      return { success: false, error: error.message };
    }
  }
  
   /**
   * Hides one or more columns.
   * @param {number} colIndex - The 0-based index of the first column to hide.
   * @param {number} count - Number of columns to hide (default: 1).
   * @returns {Promise<Object>}
   */
  async hideColumns(colIndex, count = 1) {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const rangeToHide = sheet.getRangeByIndexes(0, colIndex, 0, count).getEntireColumn();
        rangeToHide.format.columnWidth = 0;
        await context.sync();
      });
      
      this._trackOperation("hideColumns", {
        colIndex: colIndex,
        count: count
      });
      
      return { success: true };
    } catch (error) {
      console.error("Error hiding columns:", error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Unhides one or more columns.
   * @param {number} colIndex - The 0-based index of the first column to unhide.
   * @param {number} count - Number of columns to unhide (default: 1).
   * @returns {Promise<Object>}
   */
  async unhideColumns(colIndex, count = 1) {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const rangeToUnhide = sheet.getRangeByIndexes(0, colIndex, 0, count).getEntireColumn();
        rangeToUnhide.columnHidden = false;
        await context.sync();
      });
      
      this._trackOperation("unhideColumns", {
        colIndex: colIndex,
        count: count
      });
      
      return { success: true };
    } catch (error) {
      console.error("Error unhiding columns:", error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Undoes insertRows operation by deleting the inserted rows.
   * @param {Object} details - { rowIndex, count }
   * @private
   */
  async _undoInsertRows(details) {
    return this.deleteRows(details.rowIndex, details.count);
    // Note: We need to ensure deleteRows does NOT track its own operation when called from here.
    // This requires a slight modification or a flag.
    // For simplicity, the current implementation will add a 'deleteRows' to the history.
  }

  /**
   * Undoes deleteRows operation by re-inserting rows (without data for now).
   * @param {Object} details - { rowIndex, count, deletedData? }
   * @private
   */
  async _undoDeleteRows(details) {
    // Basic undo: re-insert blank rows.
    // TODO: Enhance to restore data if `details.deletedData` is implemented.
    return this.insertRows(details.rowIndex, details.count);
  }

  /**
   * Undoes insertColumns operation by deleting the inserted columns.
   * @param {Object} details - { colIndex, count }
   * @private
   */
  async _undoInsertColumns(details) {
    return this.deleteColumns(details.colIndex, details.count);
  }

  /**
   * Undoes deleteColumns operation by re-inserting columns.
   * @param {Object} details - { colIndex, count }
   * @private
   */
  async _undoDeleteColumns(details) {
    return this.insertColumns(details.colIndex, details.count);
  }

  /**
   * Undoes hideRows operation by unhiding the rows.
   * @param {Object} details - { rowIndex, count }
   * @private
   */
  async _undoHideRows(details) {
    return this.unhideRows(details.rowIndex, details.count);
  }

  /**
   * Undoes unhideRows operation by hiding the rows.
   * @param {Object} details - { rowIndex, count }
   * @private
   */
  async _undoUnhideRows(details) {
    return this.hideRows(details.rowIndex, details.count);
  }
  
  /**
   * Undoes hideColumns operation by unhiding the columns.
   * @param {Object} details - { colIndex, count }
   * @private
   */
  async _undoHideColumns(details) {
    return this.unhideColumns(details.colIndex, details.count);
  }

  /**
   * Undoes unhideColumns operation by hiding the columns.
   * @param {Object} details - { colIndex, count }
   * @private
   */
  async _undoUnhideColumns(details) {
    return this.hideColumns(details.colIndex, details.count);
  }
}

/**
 * Apply formatting only to cells that contain specific text
 * @param {string} range - Range to check (e.g., "A1:D10" or "A:A")
 * @param {string} content - Text content to match
 * @param {object} options - Formatting options (fillColor, fontColor, bold, italic, underline)
 * @returns {Promise<object>} - Result with count of formatted cells
 */
export const formatCellsByContent = async (range, content, options = {}) => {
  let formattedCount = 0;
  let affectedCellsUndoData = []; // Store previous format data for undo

  return new Promise((resolve, reject) => {
    try {
      Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        const rangeObj = worksheet.getRange(range);
        // Load only values initially needed to find matches
        rangeObj.load("address, rowCount, columnCount, values");
        await context.sync();

        console.log(`Checking range ${rangeObj.address} for cells containing "${content}"`);

        // 1. Find matching cell indices without Excel API calls inside loop
        const matchingCellIndices = [];
        for (let r = 0; r < rangeObj.rowCount; r++) {
          for (let c = 0; c < rangeObj.columnCount; c++) {
            const cellValue = rangeObj.values[r][c];
            const isMatch = cellValue !== null &&
                          cellValue !== undefined &&
                          cellValue.toString().toLowerCase().includes(content.toLowerCase());
            if (isMatch) {
              matchingCellIndices.push({ r, c });
              formattedCount++;
            }
          }
        }

        if (matchingCellIndices.length > 0) {
          // 2. Create proxy objects and prepare batch load
          const cellProxies = [];
          matchingCellIndices.forEach(({ r, c }) => {
            const cell = rangeObj.getCell(r, c);
            // 3. Load required properties for all matching cells
            cell.load("address");
            // Load specific format properties needed for undo and potential setting
            cell.format.load("fill/color, font/color, font/bold, font/italic, font/underline");
            cellProxies.push(cell);
          });

          // 4. Sync once after loading all properties
          await context.sync();

          // 5. Iterate through loaded proxies, store previous format, apply new format
          cellProxies.forEach(cell => {
            // Read loaded properties and store for undo
             affectedCellsUndoData.push({
               address: cell.address,
               previousFill: cell.format.fill.color,
               previousFontColor: cell.format.font.color,
               previousBold: cell.format.font.bold,
               previousItalic: cell.format.font.italic,
               previousUnderline: cell.format.font.underline,
             });

            // Apply formatting properties to the proxy object
            if (options.fillColor) {
              cell.format.fill.color = options.fillColor;
            }
            if (options.fontColor) {
              cell.format.font.color = options.fontColor;
            }
            if (options.bold !== undefined) {
              cell.format.font.bold = options.bold;
            }
            if (options.italic !== undefined) {
              cell.format.font.italic = options.italic;
            }
            // Check for underline property specifically
            if (options.underline !== undefined) {
              // Ensure we use the correct Excel Enum if applicable, or a valid string like "None"
              cell.format.font.underline = options.underline === true ? Excel.UnderlineStyle.single : options.underline;
            }
          });

           // 6. Sync once at the end to apply all format changes
          await context.sync();

          // Track operation for undo using the collected data
          if (excelService) {
            excelService._trackOperation("formatCellsByContent", {
              range: rangeObj.address,
              content,
              options,
              affectedCells: affectedCellsUndoData // Use the collected data
            });
          }
        } else {
           console.log("No matching cells found to format.");
        }

        resolve({
          formattedCount,
          success: true,
          message: `Applied formatting to ${formattedCount} cell(s) containing "${content}" in range ${rangeObj.address}`
        });
      }).catch(error => {
        console.error("Error in formatCellsByContent Excel.run:", error);
        // Check for the specific property not loaded error
        if (error instanceof OfficeExtension.Error && error.code === "PropertyNotLoaded") {
           reject({
             success: false,
             error: `PropertyNotLoaded: ${error.message}. Ensure properties are loaded before reading.`,
             details: error
           });
        } else {
           reject({
             success: false,
             error: error.message || "Error formatting cells",
             details: error
           });
        }
      });
    } catch (error) {
      // Catch errors thrown before Excel.run starts
      console.error("Outer Error in formatCellsByContent:", error);
      reject({
        success: false,
        error: error.message || "Error setting up formatCellsByContent",
        details: error
      });
    }
  });
};

/**
 * Format entire rows based on a condition in a specific column
 * @param {string} range - Range to check (e.g., "A1:D10" or "A:D")
 * @param {string} columnRef - Column to check for the condition (e.g., "B" or index like "2")
 * @param {string} condition - Text content to match or special keyword like "empty", "missing", "blank"
 * @param {object} options - Formatting options (fillColor, fontColor, bold, italic, underline)
 * @returns {Promise<object>} - Result with count of formatted rows
 */
export const formatRowsByCondition = async (range, columnRef, condition, options = {}) => {
  let formattedCount = 0;
  let affectedRowsUndoData = []; // Store previous format data for undo

  return new Promise((resolve, reject) => {
    try {
      Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        const rangeObj = worksheet.getRange(range);
        // Load values initially needed to find matches
        rangeObj.load("address, rowCount, columnCount, values");
        await context.sync();

        console.log(`Checking range ${rangeObj.address} for rows where column ${columnRef} satisfies condition "${condition}"`);

        // Determine which column index to check
        let columnIndex;
        if (/^\\d+$/.test(columnRef)) {
          columnIndex = parseInt(columnRef, 10) - 1; // Convert 1-based user input to 0-based index
        } else {
          // Attempt to convert column letter to 0-based index
          try {
            const tempRange = worksheet.getRange(`${columnRef}1`);
            tempRange.load("columnIndex");
            await context.sync();
            columnIndex = tempRange.columnIndex;
          } catch (e) {
             throw new Error(`Invalid column reference: ${columnRef}`);
          }
        }

        // Validate column index relative to the provided range
        const rangeStartColIndex = worksheet.getRange(range).getColumnsBefore(columnIndex).getCount();
        await context.sync();
        if (rangeStartColIndex < 0 || rangeStartColIndex >= rangeObj.columnCount) {
             throw new Error(`Column index derived from ${columnRef} is outside the specified range ${rangeObj.address}`);
        }
        const relativeColumnIndex = rangeStartColIndex; // Use the 0-based index relative to the range start

        const isEmptyCheck = ["empty", "missing", "blank", "null", "undefined"].includes(
          condition.toLowerCase().trim()
        );

        // 1. Find matching row indices without Excel API calls inside loop
        const matchingRowIndices = []; // Store 0-based row index within the range
        for (let r = 0; r < rangeObj.rowCount; r++) {
          const cellValue = rangeObj.values[r][relativeColumnIndex]; // Use relative index
          let isMatch = false;

          if (isEmptyCheck) {
            isMatch = cellValue === null ||
                     cellValue === undefined ||
                     String(cellValue).trim() === ""; // Check trimmed string value too
          } else {
            isMatch = cellValue !== null &&
                     cellValue !== undefined &&
                     String(cellValue).toLowerCase().includes(condition.toLowerCase());
          }

          if (isMatch) {
            matchingRowIndices.push(r);
            formattedCount++;
          }
        }

        if (matchingRowIndices.length > 0) {
          // 2. Create proxy objects and prepare batch load for entire rows
          const rowProxies = [];
          matchingRowIndices.forEach(rowIndex => {
            // Get the entire row object corresponding to the match within the original range
            const rowRange = rangeObj.getRow(rowIndex);
            // 3. Load required properties for all matching rows
            rowRange.load("address");
            rowRange.format.load("fill/color, font/color, font/bold, font/italic, font/underline");
            rowProxies.push(rowRange);
          });

          // 4. Sync once after loading all properties
          await context.sync();

          // 5. Iterate through loaded proxies, store previous format, apply new format
          rowProxies.forEach(rowRange => {
             // Read loaded properties and store for undo
             affectedRowsUndoData.push({
               address: rowRange.address,
               previousFill: rowRange.format.fill.color,
               previousFontColor: rowRange.format.font.color,
               previousBold: rowRange.format.font.bold,
               previousItalic: rowRange.format.font.italic,
               previousUnderline: rowRange.format.font.underline,
             });

            // Apply formatting properties to the proxy object
            if (options.fillColor) {
              rowRange.format.fill.color = options.fillColor;
            }
            if (options.fontColor) {
              rowRange.format.font.color = options.fontColor;
            }
            if (options.bold !== undefined) {
              rowRange.format.font.bold = options.bold;
            }
            if (options.italic !== undefined) {
              rowRange.format.font.italic = options.italic;
            }
            if (options.underline !== undefined) {
               rowRange.format.font.underline = options.underline === true ? Excel.UnderlineStyle.single : options.underline;
            }
          });

          // 6. Sync once at the end to apply all format changes
          await context.sync();

          // Track operation for undo using the collected data
          if (excelService) {
            excelService._trackOperation("formatRowsByCondition", {
              range: rangeObj.address,
              columnRef,
              condition,
              options,
              affectedCells: affectedRowsUndoData // Use the collected data (renamed variable for consistency)
            });
          }
        } else {
           console.log("No matching rows found to format.");
        }

        let conditionDescription = isEmptyCheck ?
          "missing or empty" :
          `containing "${condition}"`;
        resolve({
          formattedCount,
          success: true,
          message: `Applied formatting to ${formattedCount} row(s) where column ${columnRef} is ${conditionDescription} in range ${rangeObj.address}`
        });
      }).catch(error => {
        console.error("Error in formatRowsByCondition Excel.run:", error);
        if (error instanceof OfficeExtension.Error && error.code === "PropertyNotLoaded") {
             reject({
               success: false,
               error: `PropertyNotLoaded: ${error.message}. Ensure properties are loaded before reading.`,
               details: error
             });
        } else {
             reject({
               success: false,
               error: error.message || "Error formatting rows by condition",
               details: error
             });
        }
      });
    } catch (error) {
      console.error("Outer Error in formatRowsByCondition:", error);
      reject({
        success: false,
        error: error.message || "Error setting up formatRowsByCondition",
        details: error
      });
    }
  });
};

/**
 * Format rows based on exact match in a specific column (not just contains)
 * @param {string} range - Range to check (e.g., "A1:D10" or "A:D")
 * @param {string} columnRef - Column to check for the condition (e.g., "B" or index like "2")
 * @param {string} exactValue - Exact value to match (case insensitive)
 * @param {object} options - Formatting options (fillColor, fontColor, bold, italic, underline)
 * @returns {Promise<object>} - Result with count of formatted rows
 */
export const formatRowsByExactMatch = async (range, columnRef, exactValue, options = {}) => {
  let formattedCount = 0;
  let affectedRowsUndoData = []; // Store previous format data for undo

  return new Promise((resolve, reject) => {
    try {
      Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        const rangeObj = worksheet.getRange(range);
        // Load values initially needed to find matches
        rangeObj.load("address, rowCount, columnCount, values");
        await context.sync();

        console.log(`Checking range ${rangeObj.address} for rows where column ${columnRef} equals "${exactValue}"`);

        // Determine which column index to check (relative to the start of the range)
        let relativeColumnIndex;
         try {
            const fullColRange = worksheet.getRange(`${columnRef}1`);
            fullColRange.load("columnIndex");
            await context.sync();
            const absoluteColIndex = fullColRange.columnIndex;

            const rangeStartColRange = rangeObj.getColumn(0);
            rangeStartColRange.load("columnIndex");
            await context.sync();
            const rangeStartColIndex = rangeStartColRange.columnIndex;

            relativeColumnIndex = absoluteColIndex - rangeStartColIndex;

            if (relativeColumnIndex < 0 || relativeColumnIndex >= rangeObj.columnCount) {
                 throw new Error(`Column ${columnRef} is outside the specified range ${rangeObj.address}`);
            }
         } catch (e) {
             throw new Error(`Invalid column reference or range setup: ${columnRef}. ${e.message}`);
         }

        // 1. Find matching row indices without Excel API calls inside loop
        const matchingRowIndices = []; // Store 0-based row index within the range
        const lowerCaseExactValue = String(exactValue).toLowerCase(); // Prepare for comparison

        for (let r = 0; r < rangeObj.rowCount; r++) {
          const cellValue = rangeObj.values[r][relativeColumnIndex];
          // Check for exact match (case insensitive)
          const isMatch = cellValue !== null &&
                       cellValue !== undefined &&
                       String(cellValue).toLowerCase() === lowerCaseExactValue;

          if (isMatch) {
            matchingRowIndices.push(r);
            formattedCount++;
          }
        }

        if (matchingRowIndices.length > 0) {
          // 2. Create proxy objects and prepare batch load for entire rows
          const rowProxies = [];
          matchingRowIndices.forEach(rowIndex => {
            const rowRange = rangeObj.getRow(rowIndex);
            // 3. Load required properties for all matching rows
            rowRange.load("address");
            rowRange.format.load("fill/color, font/color, font/bold, font/italic, font/underline");
            rowProxies.push(rowRange);
          });

          // 4. Sync once after loading all properties
          await context.sync();

          // 5. Iterate through loaded proxies, store previous format, apply new format
          rowProxies.forEach(rowRange => {
             // Read loaded properties and store for undo
             affectedRowsUndoData.push({
               address: rowRange.address,
               previousFill: rowRange.format.fill.color,
               previousFontColor: rowRange.format.font.color,
               previousBold: rowRange.format.font.bold,
               previousItalic: rowRange.format.font.italic,
               previousUnderline: rowRange.format.font.underline,
             });

            // Apply formatting properties to the proxy object
            if (options.fillColor) {
              rowRange.format.fill.color = options.fillColor;
            }
            if (options.fontColor) {
              rowRange.format.font.color = options.fontColor;
            }
            if (options.bold !== undefined) {
              rowRange.format.font.bold = options.bold;
            }
             if (options.italic !== undefined) {
              rowRange.format.font.italic = options.italic;
            }
             if (options.underline !== undefined) {
               rowRange.format.font.underline = options.underline === true ? Excel.UnderlineStyle.single : options.underline;
            }
          });

          // 6. Sync once at the end to apply all format changes
          await context.sync();

          // Track operation for undo using the collected data
          if (excelService) {
            excelService._trackOperation("formatRowsByExactMatch", {
              range: rangeObj.address,
              columnRef,
              exactValue,
              options,
              affectedCells: affectedRowsUndoData // Use the collected data (renamed variable)
            });
          }
        } else {
           console.log("No matching rows found to format.");
        }

        resolve({
          formattedCount,
          success: true,
          message: `Applied formatting to ${formattedCount} row(s) where column ${columnRef} exactly matches "${exactValue}" in range ${rangeObj.address}`
        });
      }).catch(error => {
        console.error("Error in formatRowsByExactMatch Excel.run:", error);
        if (error instanceof OfficeExtension.Error && error.code === "PropertyNotLoaded") {
             reject({
               success: false,
               error: `PropertyNotLoaded: ${error.message}. Ensure properties are loaded before reading.`,
               details: error
             });
        } else {
             reject({
               success: false,
               error: error.message || "Error formatting rows by exact match",
               details: error
             });
        }
      });
    } catch (error) {
       console.error("Outer Error in formatRowsByExactMatch:", error);
       reject({
         success: false,
         error: error.message || "Error setting up formatRowsByExactMatch",
         details: error
       });
    }
  });
};

/**
 * Gets unique values from a specified range (e.g., a column or row)
 * @param {string} range - The range to get unique values from (e.g., "A:A" or "B2:B100")
 * @returns {Promise<{success: boolean, uniqueValues?: Array, error?: string}>}
 */
export const getUniqueValuesInRange = async (range) => {
  try {
    let uniqueValues = [];
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const rangeObj = sheet.getRange(range);
      rangeObj.load("values");
      await context.sync();
      // Flatten the 2D array and filter out empty/nulls
      const allValues = rangeObj.values.flat().filter(v => v !== null && v !== undefined && v !== "");
      uniqueValues = [...new Set(allValues)];
    });
    return { success: true, uniqueValues };
  } catch (error) {
    console.error("Error getting unique values in range:", error);
    return { success: false, error: error.message };
  }
};

// Export the service
const excelService = new ExcelService();
export default excelService; 