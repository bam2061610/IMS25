// ============================================
// Code.gs - Main Controller V2 (ИСПРАВЛЕННАЯ ВЕРСИЯ)
// ============================================

const CONFIG = {
  // ⚠️ REPLACE WITH YOUR IDs!
  SPREADSHEET_ID: '1t27lpVcDHyqIJZ-WbVVc6GXwx3fCFbI4pfvbQc5pa2Y',
  DRIVE_FOLDER_ID: '1_o-42adC6ZedX9sfiBO51zSTFSRGmrRl',
  
  // Sheet names (English)
  SHEET_NAME: 'Invoices',
  HISTORY_SHEET_NAME: 'History',
  SUPPLIERS_SHEET_NAME: 'Suppliers',
  ARCHIVE_SHEET_NAME: 'ARCH',
  
  // Settings
  MAX_FILE_SIZE: 20 * 1024 * 1024, // 20 MB
  MAX_FILES_COUNT: 3,
  ARCHIVE_DAYS: 30,
  
  // Timezone
  TIMEZONE: 'Asia/Almaty'
};

// Main function - opens web app
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Invoice Management System V2')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Include HTML files
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Get config for frontend
function getConfig() {
  return {
    companies: [
      "ТОО «Orhun Medical (Орхун Медикал)»",
      "ТОО «Orhun Lab»",
      "ТОО «ALMA MEDICAL GROUP»",
      "ТОО «Hayat Medical Group (Хаят Медикал Групп)»",
      "ТОО «4G Medtech Service (4Г Медтех Сервис)»",
      "ТОО «MARKETERA»",
      "ТОО «MedSpace Realty»",
      "ТОО «MediSupport»",
      "ТОО «Orhun Pharma»",
      "ТОО «Orhun Trade»",
      "Частная Компания «Orhun Med Limited»",
      "ТОО «Protek (Протек)»",
      "ТОО «Renova»"
    ],
    currencies: ['KZT', 'USD', 'EUR', 'RUB'],
    priorities: ['Обычный', 'Высокий', 'Срочно']
  };
}

// Test function
function testSettings() {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
    
    Logger.log('✅ Spreadsheet: ' + sheet.getName());
    Logger.log('✅ Folder: ' + folder.getName());
    
    return 'All settings are correct!';
  } catch (error) {
    Logger.log('❌ Error: ' + error.message);
    return 'Error: ' + error.message;
  }
}


// ============================================
// AUTHENTICATION & USERS API
// ============================================

/**
 * Аутентификация пользователя по коду доступа
 */
function authenticateUser(accessCode) {
  try {
    Logger.log('🔐 Authenticating user with code: ' + accessCode);
    
    const sheet = getOrCreateUsersSheet();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(accessCode)) {
        // Проверяем активность пользователя
        const isActive = data[i][4] === true || data[i][4] === 'TRUE' || data[i][4] === 'true';
        
        if (!isActive) {
          Logger.log('❌ User is deactivated');
          return {
            success: false,
            error: 'Пользователь деактивирован. Обратитесь к администратору.'
          };
        }
        
        const user = {
          code: String(data[i][0]),
          name: String(data[i][1]),
          role: String(data[i][2]),
          permissions: JSON.parse(data[i][3]),
          active: isActive
        };
        
        // Обновляем lastLogin
        const now = new Date().toLocaleString('ru-RU', {timeZone: CONFIG.TIMEZONE});
        sheet.getRange(i + 1, 7).setValue(now);
        
        Logger.log('✅ User authenticated: ' + user.name);
        return {
          success: true,
          user: user
        };
      }
    }
    
    Logger.log('❌ Authentication failed - invalid code');
    return {
      success: false,
      error: 'Неверный код доступа'
    };
    
  } catch (error) {
    Logger.log('❌ Authentication error: ' + error);
    return {
      success: false,
      error: 'Ошибка аутентификации: ' + error.message
    };
  }
}

/**
 * Получить всех пользователей (только для администратора)
 */
function getAllUsers(userInfo) {
  // Только админ может просматривать пользователей
  if (!userInfo || !userInfo.permissions || userInfo.permissions.indexOf('all') === -1) {
    return {
      success: false,
      error: 'Недостаточно прав'
    };
  }
  
  try {
    const sheet = getOrCreateUsersSheet();
    const data = sheet.getDataRange().getValues();
    
    const users = [];
    for (let i = 1; i < data.length; i++) {
      users.push({
        code: String(data[i][0]),
        name: String(data[i][1]),
        role: String(data[i][2]),
        permissions: JSON.parse(data[i][3]),
        active: data[i][4] === true || data[i][4] === 'TRUE' || data[i][4] === 'true',
        createdAt: String(data[i][5]),
        lastLogin: String(data[i][6])
      });
    }
    
    return {
      success: true,
      users: users
    };
    
  } catch (error) {
    Logger.log('❌ Error getting users: ' + error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Добавить нового пользователя (только для администратора)
 */
function addUser(userData, userInfo) {
  // Только админ может добавлять пользователей
  if (!userInfo || !userInfo.permissions || userInfo.permissions.indexOf('all') === -1) {
    return {
      success: false,
      error: 'Недостаточно прав'
    };
  }
  
  try {
    const sheet = getOrCreateUsersSheet();
    const now = new Date().toLocaleString('ru-RU', {timeZone: CONFIG.TIMEZONE});
    
    // Проверяем, не существует ли уже такой код
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(userData.code)) {
        return {
          success: false,
          error: 'Пользователь с таким кодом уже существует'
        };
      }
    }
    
    const rowData = [
      String(userData.code),
      String(userData.name),
      String(userData.role),
      JSON.stringify(userData.permissions),
      true,
      now,
      ''
    ];
    
    const nextRow = sheet.getLastRow() + 1;
    sheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
    
    Logger.log('✅ User added: ' + userData.name);
    
    // Логируем действие
    try {
      if (typeof logAction === 'function') {
        logAction(
          null,
          userInfo,
          'ADD_USER',
          '',
          '',
          'Добавлен пользователь: ' + userData.name + ' (' + userData.role + ')'
        );
      }
    } catch (logError) {
      Logger.log('⚠️ Log warning: ' + logError);
    }
    
    return {
      success: true,
      message: 'Пользователь добавлен'
    };
    
  } catch (error) {
    Logger.log('❌ Error adding user: ' + error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Деактивировать пользователя (только для администратора)
 */
function deactivateUser(accessCode, userInfo) {
  // Только админ может деактивировать пользователей
  if (!userInfo || !userInfo.permissions || userInfo.permissions.indexOf('all') === -1) {
    return {
      success: false,
      error: 'Недостаточно прав'
    };
  }
  
  // Нельзя деактивировать самого себя
  if (userInfo.code === accessCode) {
    return {
      success: false,
      error: 'Нельзя деактивировать самого себя'
    };
  }
  
  try {
    const sheet = getOrCreateUsersSheet();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(accessCode)) {
        const rowIndex = i + 1;
        const userName = data[i][1];
        
        sheet.getRange(rowIndex, 5).setValue(false);
        
        Logger.log('✅ User deactivated: ' + userName);
        
        // Логируем действие
        try {
          if (typeof logAction === 'function') {
            logAction(
              null,
              userInfo,
              'DEACTIVATE_USER',
              '',
              '',
              'Деактивирован пользователь: ' + userName
            );
          }
        } catch (logError) {
          Logger.log('⚠️ Log warning: ' + logError);
        }
        
        return {
          success: true,
          message: 'Пользователь деактивирован'
        };
      }
    }
    
    return {
      success: false,
      error: 'Пользователь не найден'
    };
    
  } catch (error) {
    Logger.log('❌ Error deactivating user: ' + error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Активировать пользователя (только для администратора)
 */
function activateUser(accessCode, userInfo) {
  // Только админ может активировать пользователей
  if (!userInfo || !userInfo.permissions || userInfo.permissions.indexOf('all') === -1) {
    return {
      success: false,
      error: 'Недостаточно прав'
    };
  }
  
  try {
    const sheet = getOrCreateUsersSheet();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(accessCode)) {
        const rowIndex = i + 1;
        const userName = data[i][1];
        
        sheet.getRange(rowIndex, 5).setValue(true);
        
        Logger.log('✅ User activated: ' + userName);
        
        // Логируем действие
        try {
          if (typeof logAction === 'function') {
            logAction(
              null,
              userInfo,
              'ACTIVATE_USER',
              '',
              '',
              'Активирован пользователь: ' + userName
            );
          }
        } catch (logError) {
          Logger.log('⚠️ Log warning: ' + logError);
        }
        
        return {
          success: true,
          message: 'Пользователь активирован'
        };
      }
    }
    
    return {
      success: false,
      error: 'Пользователь не найден'
    };
    
  } catch (error) {
    Logger.log('❌ Error activating user: ' + error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Получить или создать лист пользователей
 */
function getOrCreateUsersSheet() {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = spreadsheet.getSheetByName('ADM');
    
    if (!sheet) {
      Logger.log('📋 Creating ADM (Users) sheet');
      sheet = spreadsheet.insertSheet('ADM');
      createUsersHeaders(sheet);
    }
    
    return sheet;
  } catch (error) {
    Logger.log('❌ Error accessing Users sheet: ' + error);
    throw error;
  }
}

/**
 * Создать заголовки для листа пользователей
 */
function createUsersHeaders(sheet) {
  const headers = [
    'code',
    'name',
    'role',
    'permissions',
    'active',
    'createdAt',
    'lastLogin'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#9c27b0');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // Защита листа - только владелец может редактировать
  const protection = sheet.protect().setDescription('Users database - Admin only');
  protection.setWarningOnly(true);
}
