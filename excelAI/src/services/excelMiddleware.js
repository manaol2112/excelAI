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
  async extractDataAnalysisContext(range, options = {}) {
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
        // Enhanced header detection
        let hasHeaders = false;
        let headerWarning = null;
        if (typeof options.hasHeaders === 'boolean') {
          hasHeaders = options.hasHeaders;
        } else {
          const detection = this.advancedDetectHeaders(rangeData);
          hasHeaders = detection.hasHeaders;
          if (detection.warning) headerWarning = detection.warning;
        }
        // Extract data summary statistics
        const stats = this.extractDataStats(rangeData);
        // Add warnings if ambiguous
        const warnings = [];
        if (headerWarning) warnings.push(headerWarning);
        if (stats && stats.dataTypeWarnings && stats.dataTypeWarnings.length) {
          warnings.push(...stats.dataTypeWarnings);
        }
        return {
          success: true,
          data: rangeData,
          hasHeaders,
          stats,
          range: rangeData.address || "Unknown range",
          warnings
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
  
  // Advanced header detection with uniqueness and string ratio
  advancedDetectHeaders(rangeData) {
    if (!rangeData || !rangeData.data || rangeData.data.length === 0) {
      return { hasHeaders: false };
    }
    const firstRow = rangeData.data[0];
    const rowCount = rangeData.data.length;
    let warning = null;
    // If all strings and unique, likely headers
    const allStrings = firstRow.every(cell => typeof cell === 'string');
    const uniqueCount = new Set(firstRow).size;
    if (allStrings && uniqueCount === firstRow.length) {
      // Check if next row is mostly numbers or not strings
      if (rowCount > 1) {
        const secondRow = rangeData.data[1];
        const secondRowNumeric = secondRow.filter(cell => typeof cell === 'number' || (typeof cell === 'string' && !isNaN(parseFloat(cell)))).length;
        if (secondRowNumeric > firstRow.length / 2) {
          return { hasHeaders: true };
        }
      }
      // If only one row, ambiguous
      warning = 'Header detection ambiguous: only one row present.';
      return { hasHeaders: true, warning };
    }
    // If >60% of first row are strings and unique, likely headers
    const stringCount = firstRow.filter(cell => typeof cell === 'string').length;
    if (stringCount / firstRow.length > 0.6 && uniqueCount === firstRow.length) {
      return { hasHeaders: true };
    }
    // If not unique, warn
    if (allStrings && uniqueCount < firstRow.length) {
      warning = 'Header detection ambiguous: first row has duplicate values.';
      return { hasHeaders: false, warning };
    }
    // Fallback: not headers
    return { hasHeaders: false };
  }
  
  // Enhanced data type detection with dominant type threshold and warnings
  extractDataStats(rangeData) {
    if (!rangeData || !rangeData.data || rangeData.data.length === 0) {
      return null;
    }
    const stats = {
      rowCount: rangeData.data.length,
      columnCount: rangeData.data[0].length,
      dataTypes: [],
      emptyCells: 0,
      nonEmptyCells: 0,
      dataTypeWarnings: []
    };
    for (let colIndex = 0; colIndex < stats.columnCount; colIndex++) {
      const columnData = rangeData.data.map(row => row[colIndex]);
      const nonEmptyCount = columnData.filter(cell => cell !== null && cell !== undefined && cell !== '').length;
      // Count types
      let num = 0, date = 0, text = 0;
      for (const cell of columnData) {
        if (cell === null || cell === undefined || cell === '') continue;
        if (typeof cell === 'number' || (typeof cell === 'string' && !isNaN(parseFloat(cell)) && cell.trim() !== '')) num++;
        else if (cell instanceof Date || (typeof cell === 'string' && !isNaN(Date.parse(cell)))) date++;
        else text++;
      }
      let columnType = 'mixed';
      const maxType = Math.max(num, date, text);
      if (maxType / (nonEmptyCount || 1) > 0.6) {
        if (num === maxType) columnType = 'numeric';
        else if (date === maxType) columnType = 'date';
        else if (text === maxType) columnType = 'text';
      }
      if (columnType === 'mixed' && maxType > 0) {
        stats.dataTypeWarnings.push(`Column ${colIndex + 1} is mixed but has a dominant type: ${num >= date && num >= text ? 'numeric' : date >= text ? 'date' : 'text'}`);
      }
      stats.dataTypes.push(columnType);
      stats.emptyCells += (stats.rowCount - nonEmptyCount);
      stats.nonEmptyCells += nonEmptyCount;
    }
    return stats;
  }
}

export default new ExcelMiddleware(); 