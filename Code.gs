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
