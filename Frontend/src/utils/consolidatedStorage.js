// src/utils/consolidatedStorage.js

const STORAGE_KEY = 'consolidated_reports';

/**
 * Add reports to consolidated storage
 * @param {Array} reports - Array of report objects with {df, title, percent_cols}
 * @param {string} category - Category key (budget_results, od_vs_results, etc.)
 */
export const addReportToStorage = (reports, category) => {
  try {
    const existing = JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}');
    
    // Validate reports
    const validReports = reports.filter(report => {
      const hasTitle = report.title && typeof report.title === 'string';
      const hasData = report.df && (Array.isArray(report.df) ? report.df.length > 0 : true);
      
      if (!hasTitle || !hasData) {
        console.warn(`⚠️ Skipping invalid report:`, {
          hasTitle,
          hasData,
          title: report.title,
          dfType: typeof report.df,
          dfLength: Array.isArray(report.df) ? report.df.length : 'Not array'
        });
        return false;
      }
      
      return true;
    });
    
    if (validReports.length === 0) {
      console.warn(`⚠️ No valid reports to store for category: ${category}`);
      return;
    }
    
    existing[category] = validReports;
    localStorage.setItem(STORAGE_KEY, JSON.stringify(existing));
    
    console.log(`✅ Added ${validReports.length} valid reports to category: ${category}`);
    console.log(`🔍 Stored reports:`, validReports.map(r => ({
      title: r.title,
      dfLength: Array.isArray(r.df) ? r.df.length : 'Not array',
      percentCols: r.percent_cols
    })));
    
    // Dispatch custom event to notify other components
    window.dispatchEvent(new CustomEvent('consolidatedReportsUpdated', {
      detail: { category, reportsCount: validReports.length }
    }));
    
  } catch (error) {
    console.error('❌ Error saving reports to storage:', error);
  }
};

/**
 * Get all consolidated reports
 * @returns {Object} All stored reports by category
 */
export const getConsolidatedReports = () => {
  try {
    return JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}');
  } catch (error) {
    console.error('❌ Error reading consolidated reports:', error);
    return {};
  }
};

/**
 * Clear all consolidated reports
 */
export const clearConsolidatedReports = () => {
  try {
    localStorage.removeItem(STORAGE_KEY);
    console.log('🗑️ Cleared all consolidated reports');
    
    // Dispatch custom event
    window.dispatchEvent(new CustomEvent('consolidatedReportsCleared'));
  } catch (error) {
    console.error('❌ Error clearing consolidated reports:', error);
  }
};

/**
 * Get reports count by category
 * @returns {Object} Count of reports per category
 */
export const getReportsCount = () => {
  try {
    const reports = getConsolidatedReports();
    const counts = {};
    
    Object.keys(reports).forEach(category => {
      counts[category] = reports[category]?.length || 0;
    });
    
    return counts;
  } catch (error) {
    console.error('❌ Error getting reports count:', error);
    return {};
  }
};

/**
 * Get total number of reports across all categories
 * @returns {number} Total report count
 */
export const getTotalReportsCount = () => {
  try {
    const reports = getConsolidatedReports();
    return Object.values(reports).reduce((total, categoryReports) => {
      return total + (categoryReports?.length || 0);
    }, 0);
  } catch (error) {
    console.error('❌ Error getting total reports count:', error);
    return 0;
  }
};

/**
 * Remove reports from a specific category
 * @param {string} category - Category to clear
 */
export const clearCategoryReports = (category) => {
  try {
    const existing = getConsolidatedReports();
    delete existing[category];
    localStorage.setItem(STORAGE_KEY, JSON.stringify(existing));
    console.log(`🗑️ Cleared reports for category: ${category}`);
    
    // Dispatch custom event
    window.dispatchEvent(new CustomEvent('consolidatedReportsUpdated', {
      detail: { category, reportsCount: 0 }
    }));
  } catch (error) {
    console.error(`❌ Error clearing category ${category}:`, error);
  }
};

/**
 * Get all reports formatted for consolidated PPT generation
 * @returns {Array} Array of all reports with category info
 */
export const getAllReportsForPPT = () => {
  try {
    const reports = getConsolidatedReports();
    const allReports = [];
    
    Object.keys(reports).forEach(category => {
      const categoryReports = reports[category] || [];
      categoryReports.forEach((report) => {
        allReports.push({
          ...report,
          category
        });
      });
    });
    
    return allReports;
  } catch (error) {
    console.error('❌ Error getting all reports for PPT:', error);
    return [];
  }
};