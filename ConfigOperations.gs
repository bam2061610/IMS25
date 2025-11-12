// ============================================
// ConfigOperations.gs - Управление настройками системы
// ============================================

// Название листа с настройками
const SETTINGS_SHEET_NAME = 'Settings';

/**
 * Получить настройки из листа Settings
 */
function getSystemSettings() {
  try {
    const sheet = getOrCreateSettingsSheet();
    const data = sheet.getDataRange().getValues();
    
    const settings = {
      spreadsheetId: CONFIG.SPREADSHEET_ID,
      driveFolderId: CONFIG.DRIVE_FOLDER_ID,
      invoicesSheetName: CONFIG.SHEET_NAME,
      usersSheetName: 'ADM',
      historySheetName: CONFIG.HISTORY_SHEET_NAME,
      suppliersSheetName: CONFIG.SUPPLIERS_SHEET_NAME
    };
    
    // Читаем настройки из листа Settings
    for (let i = 1; i < data.length; i++) {
      const key = data[i][0];
      const value = data[i][1];
      
      if (key === 'DRIVE_FOLDER_ID') {
        settings.driveFolderId = value;
      } else if (key === 'INVOICES_SHEET_NAME') {
        settings.invoicesSheetName = value;
      }
    }
    
    return settings;
    
  } catch (error) {
    Logger.log('❌ Error getting settings: ' + error);
    return null;
  }
}

/**
 * Обновить настройку
 */
function updateSystemSetting(key, value, userInfo) {
  // Только админ может менять настройки
  if (!userInfo || !userInfo.permissions || userInfo.permissions.indexOf('all') === -1) {
    return {
      success: false,
      error: 'Недостаточно прав. Только администраторы могут менять настройки.'
    };
  }
  
  try {
    const sheet = getOrCreateSettingsSheet();
    const data = sheet.getDataRange().getValues();
    
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) {
      // Добавляем новую настройку
      const nextRow = sheet.getLastRow() + 1;
      sheet.getRange(nextRow, 1).setValue(key);
      sheet.getRange(nextRow, 2).setValue(value);
      sheet.getRange(nextRow, 3).setValue(new Date().toLocaleString('ru-RU', {timeZone: CONFIG.TIMEZONE}));
      sheet.getRange(nextRow, 4).setValue(userInfo.name);
    } else {
      // Обновляем существующую
      sheet.getRange(rowIndex, 2).setValue(value);
      sheet.getRange(rowIndex, 3).setValue(new Date().toLocaleString('ru-RU', {timeZone: CONFIG.TIMEZONE}));
      sheet.getRange(rowIndex, 4).setValue(userInfo.name);
    }
    
    Logger.log('✅ Setting updated: ' + key + ' = ' + value);
    
    return {
      success: true,
      message: 'Настройка обновлена'
    };
    
  } catch (error) {
    Logger.log('❌ Error updating setting: ' + error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Получить или создать лист Settings
 */
function getOrCreateSettingsSheet() {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = spreadsheet.getSheetByName(SETTINGS_SHEET_NAME);
    
    if (!sheet) {
      Logger.log('📋 Creating Settings sheet');
      sheet = spreadsheet.insertSheet(SETTINGS_SHEET_NAME);
      createSettingsHeaders(sheet);
      populateDefaultSettings(sheet);
    }
    
    return sheet;
  } catch (error) {
    Logger.log('❌ Error accessing Settings sheet: ' + error);
    throw error;
  }
}

/**
 * Создать заголовки для листа Settings
 */
function createSettingsHeaders(sheet) {
  const headers = [
    'key',
    'value',
    'updatedAt',
    'updatedBy'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#607d8b');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 400);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 150);
}

/**
 * Заполнить настройки по умолчанию
 */
function populateDefaultSettings(sheet) {
  const now = new Date().toLocaleString('ru-RU', {timeZone: CONFIG.TIMEZONE});
  
  const defaultSettings = [
    ['DRIVE_FOLDER_ID', CONFIG.DRIVE_FOLDER_ID, now, 'System'],
    ['INVOICES_SHEET_NAME', CONFIG.SHEET_NAME, now, 'System']
  ];
  
  sheet.getRange(2, 1, defaultSettings.length, 4).setValues(defaultSettings);
  Logger.log('✅ Default settings populated');
}
