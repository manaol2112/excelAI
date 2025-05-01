import excelMiddleware from './excelMiddleware';

class DataAnalysisService {
  constructor() {
    this.cache = new Map();
    this.cacheExpiry = 30000; // 30 seconds
  }

  /**
   * Generate a comprehensive data profile for a range
   * @param {string} range Optional range address, uses selected or used range if not provided
   * @returns {object} Data profile with statistics and insights
   */
  async generateDataProfile(range = null) {
    try {
      // Use the middleware to get data with context
      const rangeContext = await excelMiddleware.extractDataAnalysisContext(range);
      
      if (!rangeContext.success) {
        throw new Error(rangeContext.error || "Failed to analyze data");
      }
      
      const { data, hasHeaders, stats } = rangeContext;
      
      // Create a cache key based on range and data hash
      const cacheKey = `profile_${range || 'auto'}_${this.hashData(data.data)}`;
      
      // Check if we have a cached result
      const cachedResult = this.getFromCache(cacheKey);
      if (cachedResult) {
        return cachedResult;
      }
      
      // Get column names (use headers if available, otherwise generic names)
      const columnNames = hasHeaders 
        ? data.data[0] 
        : Array.from({ length: stats.columnCount }, (_, i) => `Column ${i + 1}`);
      
      // Data we'll analyze (skip headers row if present)
      const analysisData = hasHeaders ? data.data.slice(1) : data.data;
      
      // Initialize the profile
      const profile = {
        range: rangeContext.range,
        rowCount: hasHeaders ? stats.rowCount - 1 : stats.rowCount,
        columnCount: stats.columnCount,
        hasHeaders,
        columns: [],
        completeness: stats.nonEmptyCells / (stats.emptyCells + stats.nonEmptyCells),
        insights: []
      };
      
      // Analyze each column
      for (let colIndex = 0; colIndex < stats.columnCount; colIndex++) {
        // Extract column data
        const columnData = analysisData.map(row => row[colIndex]);
        const dataType = stats.dataTypes[colIndex];
        
        // Get column analysis based on data type
        let columnAnalysis;
        switch (dataType) {
          case 'numeric':
            columnAnalysis = this.analyzeNumericColumn(columnData);
            break;
          case 'date':
            columnAnalysis = this.analyzeDateColumn(columnData);
            break;
          case 'text':
            columnAnalysis = this.analyzeTextColumn(columnData);
            break;
          default:
            columnAnalysis = this.analyzeMixedColumn(columnData);
        }
        
        // Add column to profile
        profile.columns.push({
          name: columnNames[colIndex],
          index: colIndex,
          dataType,
          ...columnAnalysis
        });
      }
      
      // Generate overall insights
      profile.insights = this.generateInsights(profile);
      
      // Cache the result
      this.addToCache(cacheKey, profile);
      
      return {
        success: true,
        profile
      };
    } catch (error) {
      console.error("Error generating data profile:", error);
      return {
        success: false,
        error: error.message || "Unknown error analyzing data"
      };
    }
  }
  
  /**
   * Analyze a numeric column to extract statistics
   */
  analyzeNumericColumn(columnData) {
    // Filter out non-numeric and convert string numbers
    const numericValues = columnData
      .filter(cell => cell !== null && cell !== undefined && cell !== '')
      .map(cell => typeof cell === 'number' ? cell : parseFloat(cell))
      .filter(num => !isNaN(num));
    
    if (numericValues.length === 0) {
      return { empty: true };
    }
    
    // Sort values for percentile calculations
    const sortedValues = [...numericValues].sort((a, b) => a - b);
    
    // Calculate basic statistics
    const sum = numericValues.reduce((acc, val) => acc + val, 0);
    const mean = sum / numericValues.length;
    const min = sortedValues[0];
    const max = sortedValues[sortedValues.length - 1];
    const range = max - min;
    
    // Median and quartiles
    const median = this.getPercentile(sortedValues, 0.5);
    const q1 = this.getPercentile(sortedValues, 0.25);
    const q3 = this.getPercentile(sortedValues, 0.75);
    const iqr = q3 - q1;
    
    // Calculate standard deviation
    const squaredDiffs = numericValues.map(val => Math.pow(val - mean, 2));
    const variance = squaredDiffs.reduce((acc, val) => acc + val, 0) / numericValues.length;
    const stdDev = Math.sqrt(variance);
    
    // Identify possible outliers (using 1.5 * IQR rule)
    const lowerBound = q1 - 1.5 * iqr;
    const upperBound = q3 + 1.5 * iqr;
    const outliers = numericValues.filter(val => val < lowerBound || val > upperBound);
    
    // Determine if values appear to be currency
    const currencyPattern = /^\$?\s?[\d,]+(\.\d{1,2})?$/;
    const currencySamples = columnData.filter(cell => 
      typeof cell === 'string' && currencyPattern.test(cell)
    );
    const isCurrency = currencySamples.length > 0.3 * columnData.filter(cell => cell !== null && cell !== undefined && cell !== '').length;
    
    // Check for common patterns
    const isPercentage = columnData.some(cell => 
      typeof cell === 'string' && cell.includes('%')
    );
    
    return {
      count: numericValues.length,
      uniqueCount: new Set(numericValues).size,
      mean,
      median,
      min,
      max,
      range,
      stdDev,
      q1,
      q3,
      iqr,
      outlierCount: outliers.length,
      isCurrency,
      isPercentage,
      nonEmptyCount: numericValues.length,
      emptyCount: columnData.length - numericValues.length
    };
  }
  
  /**
   * Analyze a date column to extract patterns and statistics
   */
  analyzeDateColumn(columnData) {
    // Filter out non-dates and convert string dates
    const dateValues = columnData
      .filter(cell => cell !== null && cell !== undefined && cell !== '')
      .map(cell => cell instanceof Date ? cell : new Date(cell))
      .filter(date => !isNaN(date.getTime()));
    
    if (dateValues.length === 0) {
      return { empty: true };
    }
    
    // Sort dates
    const sortedDates = [...dateValues].sort((a, b) => a.getTime() - b.getTime());
    
    // Calculate statistics
    const minDate = sortedDates[0];
    const maxDate = sortedDates[sortedDates.length - 1];
    const rangeDays = (maxDate.getTime() - minDate.getTime()) / (1000 * 60 * 60 * 24);
    
    // Get day of week distribution
    const dayOfWeekCounts = [0, 0, 0, 0, 0, 0, 0]; // Sun-Sat
    dateValues.forEach(date => {
      dayOfWeekCounts[date.getDay()]++;
    });
    
    // Get month distribution
    const monthCounts = Array(12).fill(0);
    dateValues.forEach(date => {
      monthCounts[date.getMonth()]++;
    });
    
    // Get year distribution
    const yearCounts = {};
    dateValues.forEach(date => {
      const year = date.getFullYear();
      yearCounts[year] = (yearCounts[year] || 0) + 1;
    });
    
    // Check if dates are evenly spaced (e.g., monthly, weekly)
    let isEvenlySeparated = false;
    let separationType = null;
    
    if (dateValues.length > 2) {
      const timeDiffs = [];
      for (let i = 1; i < sortedDates.length; i++) {
        timeDiffs.push(sortedDates[i].getTime() - sortedDates[i-1].getTime());
      }
      
      // Check if time differences are consistent (within 5% tolerance)
      const avgDiff = timeDiffs.reduce((acc, diff) => acc + diff, 0) / timeDiffs.length;
      const isConsistent = timeDiffs.every(diff => 
        Math.abs(diff - avgDiff) / avgDiff < 0.05
      );
      
      if (isConsistent) {
        isEvenlySeparated = true;
        // Determine interval type (daily, weekly, monthly, yearly)
        const diffDays = avgDiff / (1000 * 60 * 60 * 24);
        
        if (diffDays >= 360 && diffDays <= 370) {
          separationType = 'yearly';
        } else if (diffDays >= 28 && diffDays <= 31) {
          separationType = 'monthly';
        } else if (diffDays >= 6.5 && diffDays <= 7.5) {
          separationType = 'weekly';
        } else if (diffDays >= 0.9 && diffDays <= 1.1) {
          separationType = 'daily';
        }
      }
    }
    
    return {
      count: dateValues.length,
      uniqueCount: new Set(dateValues.map(d => d.getTime())).size,
      minDate,
      maxDate,
      rangeDays,
      medianDate: sortedDates[Math.floor(sortedDates.length / 2)],
      dayOfWeekDistribution: dayOfWeekCounts,
      monthDistribution: monthCounts,
      yearDistribution: yearCounts,
      isEvenlySeparated,
      separationType,
      nonEmptyCount: dateValues.length,
      emptyCount: columnData.length - dateValues.length
    };
  }
  
  /**
   * Analyze a text column to extract patterns and statistics
   */
  analyzeTextColumn(columnData) {
    // Filter out empty cells
    const textValues = columnData.filter(cell => 
      cell !== null && cell !== undefined && cell !== ''
    );
    
    if (textValues.length === 0) {
      return { empty: true };
    }
    
    // Calculate length statistics
    const lengths = textValues.map(text => text.toString().length);
    const avgLength = lengths.reduce((acc, len) => acc + len, 0) / lengths.length;
    
    // Get unique values and their counts
    const valueCounts = {};
    textValues.forEach(value => {
      valueCounts[value] = (valueCounts[value] || 0) + 1;
    });
    
    // Sort value counts and get top 5
    const sortedValues = Object.entries(valueCounts)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5);
    
    // Check for patterns
    const patterns = this.detectTextPatterns(textValues);
    
    return {
      count: textValues.length,
      uniqueCount: Object.keys(valueCounts).length,
      avgLength,
      minLength: Math.min(...lengths),
      maxLength: Math.max(...lengths),
      mostCommonValues: sortedValues,
      patterns,
      nonEmptyCount: textValues.length,
      emptyCount: columnData.length - textValues.length
    };
  }
  
  /**
   * Analyze a mixed column to extract insights
   */
  analyzeMixedColumn(columnData) {
    // Filter out empty cells
    const nonEmptyValues = columnData.filter(cell => 
      cell !== null && cell !== undefined && cell !== ''
    );
    
    if (nonEmptyValues.length === 0) {
      return { empty: true };
    }
    
    // Count data types
    const typesCounts = {
      number: 0,
      boolean: 0,
      string: 0,
      date: 0,
      unknown: 0
    };
    
    nonEmptyValues.forEach(value => {
      if (typeof value === 'number' || (typeof value === 'string' && !isNaN(parseFloat(value)))) {
        typesCounts.number++;
      } else if (typeof value === 'boolean' || value === 'true' || value === 'false') {
        typesCounts.boolean++;
      } else if (value instanceof Date || (typeof value === 'string' && !isNaN(Date.parse(value)))) {
        typesCounts.date++;
      } else if (typeof value === 'string') {
        typesCounts.string++;
      } else {
        typesCounts.unknown++;
      }
    });
    
    // Get dominant type
    const dominantType = Object.entries(typesCounts)
      .sort((a, b) => b[1] - a[1])[0][0];
    
    // Get unique values and their counts
    const valueCounts = {};
    nonEmptyValues.forEach(value => {
      valueCounts[value] = (valueCounts[value] || 0) + 1;
    });
    
    return {
      count: nonEmptyValues.length,
      uniqueCount: Object.keys(valueCounts).size,
      typeDistribution: typesCounts,
      dominantType,
      nonEmptyCount: nonEmptyValues.length,
      emptyCount: columnData.length - nonEmptyValues.length
    };
  }
  
  /**
   * Detect common patterns in text data
   */
  detectTextPatterns(textValues) {
    const patterns = {
      isEmail: false,
      isPhoneNumber: false,
      isPostalCode: false,
      isURL: false,
      containsNumbers: false,
      isAllUppercase: false,
      isAllLowercase: false,
      isAllAlphabetic: false,
      isCode: false
    };
    
    // Regular expressions for pattern detection
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    const phoneRegex = /^[\d\+\-\(\)\s]{7,20}$/;
    const postalCodeRegex = /^[A-Z\d]{3,10}$/i;
    const urlRegex = /^(https?:\/\/)?(www\.)?[-a-zA-Z0-9@:%._\+~#=]{2,256}\.[a-z]{2,6}\b([-a-zA-Z0-9@:%_\+.~#?&//=]*)$/;
    const containsNumbersRegex = /\d/;
    const alphaRegex = /^[a-zA-Z\s]+$/;
    const codeRegex = /^[A-Z\d\-_]{3,15}$/;
    
    // Sample for pattern detection (up to 100 values)
    const sampleSize = Math.min(textValues.length, 100);
    const sample = textValues.slice(0, sampleSize);
    
    // Count matches for each pattern
    const matches = {
      email: 0,
      phone: 0,
      postal: 0,
      url: 0,
      containsNumbers: 0,
      allUpper: 0,
      allLower: 0,
      allAlpha: 0,
      code: 0
    };
    
    sample.forEach(text => {
      const stringValue = String(text);
      
      if (emailRegex.test(stringValue)) matches.email++;
      if (phoneRegex.test(stringValue)) matches.phone++;
      if (postalCodeRegex.test(stringValue)) matches.postal++;
      if (urlRegex.test(stringValue)) matches.url++;
      if (containsNumbersRegex.test(stringValue)) matches.containsNumbers++;
      if (stringValue === stringValue.toUpperCase() && stringValue !== stringValue.toLowerCase()) matches.allUpper++;
      if (stringValue === stringValue.toLowerCase() && stringValue !== stringValue.toUpperCase()) matches.allLower++;
      if (alphaRegex.test(stringValue)) matches.allAlpha++;
      if (codeRegex.test(stringValue)) matches.code++;
    });
    
    // A pattern is identified if at least 80% of the sample matches
    const threshold = 0.8 * sampleSize;
    
    patterns.isEmail = matches.email >= threshold;
    patterns.isPhoneNumber = matches.phone >= threshold;
    patterns.isPostalCode = matches.postal >= threshold;
    patterns.isURL = matches.url >= threshold;
    patterns.containsNumbers = matches.containsNumbers >= threshold;
    patterns.isAllUppercase = matches.allUpper >= threshold;
    patterns.isAllLowercase = matches.allLower >= threshold;
    patterns.isAllAlphabetic = matches.allAlpha >= threshold;
    patterns.isCode = matches.code >= threshold;
    
    return patterns;
  }
  
  /**
   * Generate overall insights based on the data profile
   */
  generateInsights(profile) {
    const insights = [];
    
    // Check data completeness
    if (profile.completeness < 0.8) {
      insights.push({
        type: 'data_quality',
        message: `Data is ${Math.round(profile.completeness * 100)}% complete. Consider addressing missing values.`
      });
    }
    
    // Identify columns with outliers
    profile.columns.forEach(column => {
      if (column.dataType === 'numeric' && column.outlierCount > 0) {
        const outlierPercent = (column.outlierCount / column.count * 100).toFixed(1);
        insights.push({
          type: 'outliers',
          message: `Column "${column.name}" has ${column.outlierCount} outliers (${outlierPercent}%).`
        });
      }
    });
    
    // Identify potential key columns
    const potentialKeyColumns = profile.columns.filter(column => 
      column.uniqueCount === column.nonEmptyCount && column.nonEmptyCount > 0
    );
    
    if (potentialKeyColumns.length > 0) {
      potentialKeyColumns.forEach(column => {
        insights.push({
          type: 'key_column',
          message: `Column "${column.name}" could be a key column (all values are unique).`
        });
      });
    }
    
    // Identify date patterns
    const dateColumns = profile.columns.filter(column => column.dataType === 'date');
    dateColumns.forEach(column => {
      if (column.isEvenlySeparated && column.separationType) {
        insights.push({
          type: 'time_series',
          message: `Column "${column.name}" contains ${column.separationType} data, suitable for time series analysis.`
        });
      }
    });
    
    // Identify categorical columns
    profile.columns.forEach(column => {
      if (column.dataType === 'text' && 
          column.uniqueCount > 1 && 
          column.uniqueCount <= 20 && 
          column.uniqueCount / column.count <= 0.2) {
        insights.push({
          type: 'categorical',
          message: `Column "${column.name}" appears to be categorical with ${column.uniqueCount} categories.`
        });
      }
    });
    
    return insights;
  }
  
  /**
   * Calculate percentile from sorted array
   */
  getPercentile(sortedArray, percentile) {
    if (sortedArray.length === 0) return null;
    
    const index = percentile * (sortedArray.length - 1);
    const lower = Math.floor(index);
    const upper = Math.ceil(index);
    const weight = index % 1;
    
    if (upper >= sortedArray.length) return sortedArray[sortedArray.length - 1];
    if (lower === upper) return sortedArray[lower];
    
    return (1 - weight) * sortedArray[lower] + weight * sortedArray[upper];
  }
  
  /**
   * Create a simple hash of data for caching purposes
   */
  hashData(data) {
    if (!data || !data.length) return 'empty';
    
    let hash = 0;
    const str = JSON.stringify(data);
    
    for (let i = 0; i < str.length; i++) {
      const char = str.charCodeAt(i);
      hash = ((hash << 5) - hash) + char;
      hash = hash & hash; // Convert to 32bit integer
    }
    
    return hash.toString(36);
  }
  
  /**
   * Add item to cache with expiry
   */
  addToCache(key, value) {
    const item = {
      value,
      timestamp: Date.now()
    };
    this.cache.set(key, item);
    
    // Clear expired items occasionally
    if (Math.random() < 0.1) {
      this.clearExpiredCache();
    }
  }
  
  /**
   * Get item from cache if not expired
   */
  getFromCache(key) {
    const item = this.cache.get(key);
    
    if (!item) return null;
    
    // Check if expired
    if (Date.now() - item.timestamp > this.cacheExpiry) {
      this.cache.delete(key);
      return null;
    }
    
    return item.value;
  }
  
  /**
   * Clear expired items from cache
   */
  clearExpiredCache() {
    const now = Date.now();
    
    for (const [key, item] of this.cache.entries()) {
      if (now - item.timestamp > this.cacheExpiry) {
        this.cache.delete(key);
      }
    }
  }
}

export default new DataAnalysisService(); 