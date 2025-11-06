// ============================================
// LogOperations.gs - Work with History sheet (ИСПРАВЛЕННАЯ ВЕРСИЯ)
// ============================================

// Log user action
function logAction(invoiceId, userInfo, action, oldStatus, newStatus, details) {
  try {
    const sheet = getOrCreateHistorySheet();
    
    const invoice = invoiceId ? getInvoiceById(invoiceId) : null;
    const invoiceNumber = invoice ? invoice.number : '';
    
    const timestamp = new Date().toLocaleString('ru-RU', {timeZone: CONFIG.TIMEZONE || 'Asia/Almaty'});
    
    const rowData = [
      timestamp,
      userInfo.name || '',
      userInfo.role || '',
      action,
      invoiceNumber,
      oldStatus || '',
      newStatus || '',
      details || ''
    ];
    
    const nextRow = sheet.getLastRow() + 1;
    sheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
    
    Logger.log('📝 Action logged: ' + action);
    
  } catch (error) {
    Logger.log('❌ Error logging action: ' + error);
    // Не прерываем выполнение, если логирование не удалось
  }
}

// Get logs (only for admin)
function getLogs(filters) {
  try {
    const sheet = getOrCreateHistorySheet();
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return [];
    }
    
    const logs = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      const log = {
        timestamp: String(row[0] || ''),
        user: String(row[1] || ''),
        role: String(row[2] || ''),
        action: String(row[3] || ''),
        invoiceNumber: String(row[4] || ''),
        oldStatus: String(row[5] || ''),
        newStatus: String(row[6] || ''),
        details: String(row[7] || '')
      };
      
      // Apply filters if provided
      if (filters) {
        if (filters.user && log.user !== filters.user) continue;
        if (filters.action && log.action !== filters.action) continue;
        if (filters.invoiceNumber && log.invoiceNumber !== filters.invoiceNumber) continue;
        if (filters.dateFrom) {
          const logDate = new Date(log.timestamp);
          const filterDate = new Date(filters.dateFrom);
          if (logDate < filterDate) continue;
        }
        if (filters.dateTo) {
          const logDate = new Date(log.timestamp);
          const filterDate = new Date(filters.dateTo);
          if (logDate > filterDate) continue;
        }
      }
      
      logs.push(log);
    }
    
    Logger.log('📋 Loaded ' + logs.length + ' logs');
    return logs;
    
  } catch (error) {
    Logger.log('❌ Error loading logs: ' + error);
    return [];
  }
}

// Get logs for specific invoice
function getInvoiceLogs(invoiceNumber) {
  try {
    return getLogs({ invoiceNumber: invoiceNumber });
  } catch (error) {
    Logger.log('❌ Error loading invoice logs: ' + error);
    return [];
  }
}

// Get logs for specific user
function getUserLogs(userName, limit) {
  try {
    const allLogs = getLogs({ user: userName });
    
    if (limit && limit > 0) {
      return allLogs.slice(0, limit);
    }
    
    return allLogs;
  } catch (error) {
    Logger.log('❌ Error loading user logs: ' + error);
    return [];
  }
}

// Clear old logs (for maintenance)
function clearOldLogs(daysToKeep) {
  try {
    if (!daysToKeep || daysToKeep < 1) {
      throw new Error('Invalid daysToKeep parameter');
    }
    
    const sheet = getOrCreateHistorySheet();
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      Logger.log('No logs to clear');
      return { success: true, deleted: 0 };
    }
    
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - daysToKeep);
    
    let deletedCount = 0;
    
    // Iterate from bottom to top to avoid index issues
    for (let i = data.length - 1; i >= 1; i--) {
      const logDate = new Date(data[i][0]);
      
      if (logDate < cutoffDate) {
        sheet.deleteRow(i + 1);
        deletedCount++;
      }
    }
    
    Logger.log('🗑️ Cleared ' + deletedCount + ' old logs');
    return { success: true, deleted: deletedCount };
    
  } catch (error) {
    Logger.log('❌ Error clearing old logs: ' + error);
    return { success: false, error: error.message };
  }
}

// ============================================
// HELPER FUNCTIONS
// ============================================

// Get or create History sheet
function getOrCreateHistorySheet() {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = spreadsheet.getSheetByName(CONFIG.HISTORY_SHEET_NAME);
    
    if (!sheet) {
      Logger.log('📋 Creating History sheet');
      sheet = spreadsheet.insertSheet(CONFIG.HISTORY_SHEET_NAME);
      createHistoryHeaders(sheet);
    }
    
    return sheet;
  } catch (error) {
    Logger.log('❌ Error accessing History sheet: ' + error);
    throw error;
  }
}

// Create History headers
function createHistoryHeaders(sheet) {
  const headers = [
    'timestamp',
    'user',
    'role',
    'action',
    'invoiceNumber',
    'oldStatus',
    'newStatus',
    'details'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#ea4335');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // Set column widths for better readability
  sheet.setColumnWidth(1, 150); // timestamp
  sheet.setColumnWidth(2, 150); // user
  sheet.setColumnWidth(3, 100); // role
  sheet.setColumnWidth(4, 150); // action
  sheet.setColumnWidth(5, 120); // invoiceNumber
  sheet.setColumnWidth(6, 120); // oldStatus
  sheet.setColumnWidth(7, 120); // newStatus
  sheet.setColumnWidth(8, 300); // details
}
