// ============================================
// SupplierOperations.gs - Work with Suppliers sheet (ИСПРАВЛЕННАЯ ВЕРСИЯ)
// ============================================

// Get all suppliers (for autocomplete)
function getSuppliers() {
  try {
    const sheet = getOrCreateSuppliersSheet();
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return [];
    }
    
    const suppliers = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      suppliers.push({
        name: String(row[0] || ''),
        bin: String(row[1] || ''),
        count: parseInt(row[2]) || 0,
        lastUsed: String(row[3] || '')
      });
    }
    
    // Sort by usage count (most used first)
    suppliers.sort((a, b) => b.count - a.count);
    
    Logger.log('📋 Loaded ' + suppliers.length + ' suppliers');
    return suppliers;
    
  } catch (error) {
    Logger.log('❌ Error loading suppliers: ' + error);
    return [];
  }
}

// Update or create supplier template
function updateSupplierTemplate(name, bin) {
  try {
    if (!name) {
      Logger.log('⚠️ Supplier name is empty, skipping update');
      return;
    }
    
    const sheet = getOrCreateSuppliersSheet();
    const data = sheet.getDataRange().getValues();
    
    let found = false;
    let rowIndex = -1;
    
    // Search for existing supplier (case-insensitive)
    const normalizedName = String(name).trim().toLowerCase();
    for (let i = 1; i < data.length; i++) {
      const existingName = String(data[i][0]).trim().toLowerCase();
      if (existingName === normalizedName) {
        found = true;
        rowIndex = i + 1;
        break;
      }
    }
    
    const now = new Date().toLocaleString('ru-RU', {timeZone: CONFIG.TIMEZONE || 'Asia/Almaty'});
    
    if (found) {
      // Update existing
      const currentCount = parseInt(data[rowIndex-1][2]) || 0;
      sheet.getRange(rowIndex, 3).setValue(currentCount + 1); // Increase count
      sheet.getRange(rowIndex, 4).setValue(now); // Update lastUsed
      
      // Update BIN if provided and different
      if (bin && bin !== data[rowIndex-1][1]) {
        sheet.getRange(rowIndex, 2).setValue(String(bin).trim());
      }
      
      Logger.log('✅ Supplier updated: ' + name);
      
    } else {
      // Create new
      const rowData = [
        String(name).trim(),
        bin ? String(bin).trim() : '',
        1,
        now
      ];
      
      const nextRow = sheet.getLastRow() + 1;
      sheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
      
      Logger.log('✅ Supplier created: ' + name);
    }
    
  } catch (error) {
    Logger.log('❌ Error updating supplier: ' + error);
    // Не прерываем выполнение, если обновление поставщика не удалось
  }
}

// Get supplier by name
function getSupplierByName(name) {
  try {
    if (!name) return null;
    
    const sheet = getOrCreateSuppliersSheet();
    const data = sheet.getDataRange().getValues();
    
    const normalizedName = String(name).trim().toLowerCase();
    
    for (let i = 1; i < data.length; i++) {
      const existingName = String(data[i][0]).trim().toLowerCase();
      if (existingName === normalizedName) {
        return {
          name: String(data[i][0]),
          bin: String(data[i][1] || ''),
          count: parseInt(data[i][2]) || 0,
          lastUsed: String(data[i][3] || '')
        };
      }
    }
    
    return null;
  } catch (error) {
    Logger.log('❌ Error getting supplier: ' + error);
    return null;
  }
}

// Delete supplier
function deleteSupplier(name) {
  try {
    if (!name) {
      throw new Error('Supplier name is required');
    }
    
    const sheet = getOrCreateSuppliersSheet();
    const data = sheet.getDataRange().getValues();
    
    const normalizedName = String(name).trim().toLowerCase();
    
    for (let i = 1; i < data.length; i++) {
      const existingName = String(data[i][0]).trim().toLowerCase();
      if (existingName === normalizedName) {
        sheet.deleteRow(i + 1);
        Logger.log('🗑️ Supplier deleted: ' + name);
        return { success: true };
      }
    }
    
    return { success: false, error: 'Supplier not found' };
    
  } catch (error) {
    Logger.log('❌ Error deleting supplier: ' + error);
    return { success: false, error: error.message };
  }
}

// Get most used suppliers
function getMostUsedSuppliers(limit) {
  try {
    const suppliers = getSuppliers();
    
    if (!limit || limit < 1) {
      limit = 10;
    }
    
    // Already sorted by count in getSuppliers()
    return suppliers.slice(0, limit);
    
  } catch (error) {
    Logger.log('❌ Error getting most used suppliers: ' + error);
    return [];
  }
}

// ============================================
// HELPER FUNCTIONS
// ============================================

// Get or create Suppliers sheet
function getOrCreateSuppliersSheet() {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = spreadsheet.getSheetByName(CONFIG.SUPPLIERS_SHEET_NAME);
    
    if (!sheet) {
      Logger.log('📋 Creating Suppliers sheet');
      sheet = spreadsheet.insertSheet(CONFIG.SUPPLIERS_SHEET_NAME);
      createSuppliersHeaders(sheet);
    }
    
    return sheet;
  } catch (error) {
    Logger.log('❌ Error accessing Suppliers sheet: ' + error);
    throw error;
  }
}

// Create Suppliers headers
function createSuppliersHeaders(sheet) {
  const headers = [
    'name',
    'bin',
    'count',
    'lastUsed'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#34a853');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // Set column widths for better readability
  sheet.setColumnWidth(1, 300); // name
  sheet.setColumnWidth(2, 120); // bin
  sheet.setColumnWidth(3, 80);  // count
  sheet.setColumnWidth(4, 150); // lastUsed
}
