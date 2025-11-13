// ============================================
// ArchiveOperations.gs - Система архивирования
// ============================================

/**
 * Получить лист ARCH
 */
function getArchiveSheet() {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(CONFIG.ARCHIVE_SHEET_NAME);
    
    if (!sheet) {
      throw new Error('Лист ARCH не найден! Создайте его вручную.');
    }
    
    return sheet;
  } catch (error) {
    Logger.log('❌ Error accessing ARCH sheet: ' + error);
    throw error;
  }
}

/**
 * Перенести счет в архив (вручную)
 * @param {number} invoiceId - ID счета
 * @param {object} userInfo - Информация о пользователе
 * @returns {object} Результат операции
 */
function moveInvoiceToArchive(invoiceId, userInfo) {
  // Только админ может архивировать
  if (!userInfo || !userInfo.permissions || userInfo.permissions.indexOf('all') === -1) {
    return {
      success: false,
      error: 'Недостаточно прав. Только администраторы могут архивировать счета.'
    };
  }
  
  const lock = LockService.getScriptLock();
  
  try {
    lock.waitLock(30000);
    
    Logger.log('📦 Moving invoice ' + invoiceId + ' to archive');
    
    const invoicesSheet = getOrCreateSheet();
    const archiveSheet = getArchiveSheet();
    const data = invoicesSheet.getDataRange().getValues();
    
    // Найти строку со счетом
    let rowIndex = -1;
    let rowData = null;
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(invoiceId)) {
        rowIndex = i + 1; // +1 потому что массив с 0, а строки с 1
        rowData = data[i];
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, error: 'Счет не найден' };
    }
    
    // Проверка: можно архивировать только paid или rejected
    const status = rowData[11]; // колонка L (status)
    if (status !== 'paid' && status !== 'rejected') {
      return {
        success: false,
        error: 'Можно архивировать только оплаченные или отклоненные счета'
      };
    }
    
    // Добавить строку в ARCH
    const nextArchiveRow = archiveSheet.getLastRow() + 1;
    archiveSheet.getRange(nextArchiveRow, 1, 1, rowData.length).setValues([rowData]);
    
    // Удалить строку из Invoices
    invoicesSheet.deleteRow(rowIndex);
    
    // Логировать действие
    try {
      if (typeof logAction === 'function') {
        logAction(
          invoiceId,
          userInfo,
          'MOVE_TO_ARCHIVE',
          status,
          status,
          'Счет №' + rowData[1] + ' перенесен в архив вручную'
        );
      }
    } catch (logError) {
      Logger.log('⚠️ Log warning: ' + logError);
    }
    
    Logger.log('✅ Invoice moved to archive successfully');
    return {
      success: true,
      message: 'Счет успешно архивирован'
    };
    
  } catch (error) {
    Logger.log('❌ Error moving to archive: ' + error);
    return {
      success: false,
      error: error.message || 'Не удалось архивировать счет'
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Автоматическое архивирование старых счетов
 * Вызывается триггером каждую субботу в 05:00
 */
function autoArchiveOldInvoices() {
  const lock = LockService.getScriptLock();
  
  try {
    lock.waitLock(30000);
    
    Logger.log('🤖 Starting automatic archiving...');
    
    const invoicesSheet = getOrCreateSheet();
    const archiveSheet = getArchiveSheet();
    const data = invoicesSheet.getDataRange().getValues();
    
    const now = new Date();
    const tenDaysAgo = new Date(now.getTime() - (10 * 24 * 60 * 60 * 1000));
    
    let movedCount = 0;
    const rowsToDelete = []; // Индексы строк для удаления (в обратном порядке)
    
    // Проходим по всем счетам (пропускаем заголовок)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const invoiceId = row[0];
      const status = row[11]; // колонка L (status)
      
      // Пропускаем, если статус не paid и не rejected
      if (status !== 'paid' && status !== 'rejected') {
        continue;
      }
      
      let actionDate = null;
      
      // Определяем дату действия
      if (status === 'paid') {
        actionDate = row[23]; // колонка X (paidAt)
      } else if (status === 'rejected') {
        actionDate = row[15]; // колонка P (approvedAt)
      }
      
      // Пропускаем, если дата не указана
      if (!actionDate) {
        continue;
      }
      
      // Парсим дату
      let parsedDate;
      try {
        parsedDate = new Date(actionDate);
      } catch (e) {
        Logger.log('⚠️ Cannot parse date for invoice ' + invoiceId + ': ' + actionDate);
        continue;
      }
      
      // Проверяем, прошло ли 10 дней
      if (parsedDate < tenDaysAgo) {
        // Добавить строку в ARCH
        const nextArchiveRow = archiveSheet.getLastRow() + 1;
        archiveSheet.getRange(nextArchiveRow, 1, 1, row.length).setValues([row]);
        
        // Запомнить индекс для удаления
        rowsToDelete.push(i + 1); // +1 потому что строки начинаются с 1
        
        movedCount++;
        Logger.log('📦 Archived invoice #' + invoiceId + ' (' + status + ', ' + actionDate + ')');
      }
    }
    
    // Удаляем строки в обратном порядке (чтобы индексы не сбивались)
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      invoicesSheet.deleteRow(rowsToDelete[i]);
    }
    
    // Логировать результат
    try {
      if (typeof logAction === 'function') {
        logAction(
          null,
          { name: 'System', role: 'Автоматизация' },
          'AUTO_ARCHIVE',
          '',
          '',
          'Автоархивирование: перенесено ' + movedCount + ' счетов'
        );
      }
    } catch (logError) {
      Logger.log('⚠️ Log warning: ' + logError);
    }
    
    Logger.log('✅ Automatic archiving completed: ' + movedCount + ' invoices moved');
    
    return {
      success: true,
      movedCount: movedCount
    };
    
  } catch (error) {
    Logger.log('❌ Error in automatic archiving: ' + error);
    return {
      success: false,
      error: error.message
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Получить архивные счета
 * @returns {array} Массив архивных счетов
 */
function getArchivedInvoices() {
  try {
    Logger.log('📊 Loading archived invoices...');
    
    const sheet = getArchiveSheet();
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      Logger.log('📊 Archive is empty');
      return [];
    }
    
    const invoices = [];
    const now = new Date();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      if (!row[0]) continue;
      
      // Подсчет дней в системе (фиксированный - от создания до архивирования)
      const createdDate = new Date(row[13]); // createdAt
      let daysInSystem = 0;
      
      if (row[11] === 'paid') {
        const paidDate = new Date(row[23]);
        if (!isNaN(paidDate.getTime()) && !isNaN(createdDate.getTime())) {
          daysInSystem = Math.floor((paidDate - createdDate) / (1000 * 60 * 60 * 24));
        }
      } else if (row[11] === 'rejected') {
        const rejectedDate = new Date(row[15]);
        if (!isNaN(rejectedDate.getTime()) && !isNaN(createdDate.getTime())) {
          daysInSystem = Math.floor((rejectedDate - createdDate) / (1000 * 60 * 60 * 24));
        }
      }
      
      const invoice = {
        id: row[0] ? parseInt(row[0]) : 0,
        number: String(row[1] || ''),
        date: formatDateForJS(row[2]),
        company: String(row[3] || ''),
        supplier: String(row[4] || ''),
        supplierBIN: String(row[5] || ''),
        amount: parseFloat(row[6]) || 0,
        currency: String(row[7] || 'KZT'),
        purpose: String(row[8] || ''),
        dueDate: formatDateForJS(row[9]),
        priority: String(row[10] || 'Обычный'),
        status: String(row[11] || 'pending'),
        createdBy: String(row[12] || ''),
        createdAt: String(row[13] || ''),
        approvedBy: String(row[14] || ''),
        approvedAt: String(row[15] || ''),
        confirmedBy1: String(row[16] || ''),
        confirmedAt1: String(row[17] || ''),
        confirmedBy2: String(row[18] || ''),
        confirmedAt2: String(row[19] || ''),
        paidBy: String(row[22] || ''),
        paidAt: String(row[23] || ''),
        notes: String(row[24] || ''),
        files: String(row[25] || ''),
        printed: row[26] === true || row[26] === 'TRUE' || row[26] === 'true',
        printedBy: String(row[27] || ''),
        printedAt: String(row[28] || ''),
        archived: true, // Всегда true для архивных
        comments: parseComments(row[30]),
        daysInSystem: daysInSystem
      };
      
      if (invoice.status === 'rejected') {
        invoice.rejectedBy = invoice.approvedBy;
        invoice.rejectedAt = invoice.approvedAt;
        
        if (invoice.notes) {
          const rejectionMatch = invoice.notes.match(/\[ОТКЛОНЕН[^\]]*\]:\s*([^\n]+)/);
          if (rejectionMatch) {
            invoice.rejectionReason = rejectionMatch[1];
          }
        }
      }
      
      invoices.push(invoice);
    }
    
    Logger.log('✅ Loaded ' + invoices.length + ' archived invoices');
    return invoices;
    
  } catch (error) {
    Logger.log('❌ Error loading archived invoices: ' + error);
    return [];
  }
}

/**
 * Форматирование даты для JavaScript (helper)
 */
function formatDateForJS(dateValue) {
  if (!dateValue) return '';
  
  try {
    let date;
    if (dateValue instanceof Date) {
      date = dateValue;
    } else {
      date = new Date(dateValue);
    }
    
    if (isNaN(date.getTime())) {
      return String(dateValue);
    }
    
    return date.getFullYear() + '-' + 
           String(date.getMonth() + 1).padStart(2, '0') + '-' + 
           String(date.getDate()).padStart(2, '0');
  } catch (error) {
    return String(dateValue);
  }
}

/**
 * Парсинг комментариев (helper)
 */
function parseComments(commentsStr) {
  if (!commentsStr || commentsStr === '') return [];
  
  try {
    return JSON.parse(commentsStr);
  } catch (error) {
    return [];
  }
}
