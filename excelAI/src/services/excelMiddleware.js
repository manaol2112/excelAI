import excelService from './excelService';

class ExcelMiddleware {
  constructor() {
    this.pendingOperations = new Map();
    this.operationQueue = [];
    this.isProcessing = false;
  }

  // Queue operations to prevent race conditions
  async executeOperation(operationFn, ...args) {
    const operationId = Date.now() + Math.random().toString(36).substring(2, 9);
    
    return new Promise((resolve, reject) => {
      this.operationQueue.push({
        id: operationId,
        operation: async () => {
          try {
            const result = await operationFn(...args);
            resolve(result);
          } catch (error) {
            console.error(`Operation error (${operationId}):`, error);
            reject(error);
          } finally {
            this.pendingOperations.delete(operationId);
            this.processNextOperation();
          }
        }
      });
      
      this.pendingOperations.set(operationId, { resolve, reject });
      
      if (!this.isProcessing) {
        this.processNextOperation();
      }
    });
  }
  
  async processNextOperation() {
    if (this.operationQueue.length === 0) {
      this.isProcessing = false;
      return;
    }
    
    this.isProcessing = true;
    const nextOperation = this.operationQueue.shift();
    await nextOperation.operation();
  }
  
  // Enhanced data loading with retry logic and validation
  async getWorkbookState(forceRefresh = false) {
    return this.executeOperation(async () => {
      try {
        const workbookData = await excelService.preloadWorkbookData(forceRefresh);
        
        // Validate workbook data
        if (!workbookData || !workbookData.activeWorksheet) {
          throw new Error("Failed to load complete workbook data");
        }
        
        return {
          success: true,
          data: workbookData
        };
      } catch (error) {
        console.error("Error loading workbook state:", error);
        
        // Retry once with force refresh
        if (!forceRefresh) {
          console.log("Retrying with force refresh...");
          return this.getWorkbookState(true);
        }
        
        return {
          success: false,
          error: error.message || "Unknown error loading workbook state"
        };
      }
    });
  }
  
  // Improved data analysis extraction
  async extractDataAnalysisContext(range) {
    return this.executeOperation(async () => {
      try {
        // Get data for the specified range
        let rangeData;
        if (range) {
          rangeData = await excelService.getData(range);
        } else {
          // Try to get selected range first
          const selectedRange = await excelService.getSelectedRange();
          
          if (selectedRange && selectedRange.success && selectedRange.data.address) {
            rangeData = selectedRange;
          } else {
            // Fall back to used range
            rangeData = await excelService.getAllData();
          }
        }
        
        if (!rangeData || !rangeData.success) {
          throw new Error("Failed to extract data for analysis");
        }
        
        // Get headers if available
        const hasHeaders = this.detectHeaders(rangeData);
        
        // Extract data summary statistics
        const stats = this.extractDataStats(rangeData);
        
        return {
          success: true,
          data: rangeData,
          hasHeaders,
          stats,
          range: rangeData.address || "Unknown range"
        };
      } catch (error) {
        console.error("Error extracting data analysis context:", error);
        return {
          success: false,
          error: error.message || "Unknown error analyzing data"
        };
      }
    });
  }
  
  // Detect if data has headers based on data patterns
  detectHeaders(rangeData) {
    if (!rangeData || !rangeData.data || rangeData.data.length === 0) {
      return false;
    }
    
    const firstRow = rangeData.data[0];
    
    // Check if first row contains all strings while other rows have numeric values
    const firstRowAllStrings = firstRow.every(cell => typeof cell === 'string');
    
    if (firstRowAllStrings && rangeData.data.length > 1) {
      const secondRow = rangeData.data[1];
      const secondRowHasNumbers = secondRow.some(cell => 
        typeof cell === 'number' || (typeof cell === 'string' && !isNaN(parseFloat(cell)))
      );
      
      if (secondRowHasNumbers) {
        return true;
      }
    }
    
    // Check for special header formatting (could be enhanced with formatting info)
    return false;
  }
  
  // Extract basic statistics about the data
  extractDataStats(rangeData) {
    if (!rangeData || !rangeData.data || rangeData.data.length === 0) {
      return null;
    }
    
    const stats = {
      rowCount: rangeData.data.length,
      columnCount: rangeData.data[0].length,
      dataTypes: [],
      emptyCells: 0,
      nonEmptyCells: 0
    };
    
    // Analyze each column
    for (let colIndex = 0; colIndex < stats.columnCount; colIndex++) {
      const columnData = rangeData.data.map(row => row[colIndex]);
      
      // Count non-empty cells in this column
      const nonEmptyCount = columnData.filter(cell => 
        cell !== null && cell !== undefined && cell !== ''
      ).length;
      
      // Detect data type for this column
      let columnType = 'mixed';
      
      const numericCount = columnData.filter(cell => 
        typeof cell === 'number' || (typeof cell === 'string' && !isNaN(parseFloat(cell)) && cell.trim() !== '')
      ).length;
      
      const dateCount = columnData.filter(cell =>
        cell instanceof Date || (typeof cell === 'string' && !isNaN(Date.parse(cell)))
      ).length;
      
      if (numericCount > 0.7 * nonEmptyCount) {
        columnType = 'numeric';
      } else if (dateCount > 0.7 * nonEmptyCount) {
        columnType = 'date';
      } else if (nonEmptyCount > 0) {
        columnType = 'text';
      }
      
      stats.dataTypes.push(columnType);
      stats.emptyCells += (stats.rowCount - nonEmptyCount);
      stats.nonEmptyCells += nonEmptyCount;
    }
    
    return stats;
  }
}

export default new ExcelMiddleware(); 