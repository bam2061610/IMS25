// ============================================
// SheetOperations.gs - Work with Invoices sheet (ИСПРАВЛЕННАЯ ВЕРСИЯ)
// ============================================

// Get all invoices from sheet
function getInvoices() {
  try {
    Logger.log('📊 Loading invoices from sheet...');
    
    const sheet = getOrCreateSheet();
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      Logger.log('📊 Sheet is empty');
      return [];
    }
    
    const invoices = [];
    const now = new Date();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      if (!row[0]) continue;
      
      // ПРАВИЛЬНЫЙ подсчет дней в системе
      const createdDate = new Date(row[13]); // createdAt
      let daysInSystem = 0;
      
      if (row[11] === 'paid' || row[11] === 'rejected') {
        // Для оплаченных/отклоненных - ФИКСИРОВАННОЕ количество дней
        const endDate = row[11] === 'paid' ? new Date(row[23]) : new Date(row[15]);
        if (!isNaN(endDate.getTime()) && !isNaN(createdDate.getTime())) {
          daysInSystem = Math.floor((endDate - createdDate) / (1000 * 60 * 60 * 24));
        }
      } else {
        // Для активных счетов - считаем до СЕЙЧАС
        if (!isNaN(createdDate.getTime())) {
          daysInSystem = Math.floor((now - createdDate) / (1000 * 60 * 60 * 24));
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
        archived: row[29] === true || row[29] === 'TRUE' || row[29] === 'true',
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
    
    Logger.log('✅ Loaded ' + invoices.length + ' invoices');
    return invoices;
    
  } catch (error) {
    Logger.log('❌ Error loading invoices: ' + error);
    throw error;
  }
}

// Get single invoice by ID
function getInvoiceById(invoiceId) {
  try {
    const sheet = getOrCreateSheet();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(invoiceId)) {
        const row = data[i];
        return {
          id: parseInt(row[0]),
          number: String(row[1]),
          date: formatDateForJS(row[2]),
          company: String(row[3]),
          supplier: String(row[4]),
          supplierBIN: String(row[5]),
          amount: parseFloat(row[6]),
          currency: String(row[7]),
          purpose: String(row[8]),
          dueDate: formatDateForJS(row[9]),
          priority: String(row[10]),
          status: String(row[11]),
          createdBy: String(row[12]),
          createdAt: String(row[13]),
          approvedBy: String(row[14]),
          approvedAt: String(row[15]),
          confirmedBy1: String(row[16]),
          confirmedAt1: String(row[17]),
          confirmedBy2: String(row[18]),
          confirmedAt2: String(row[19]),
          confirmedBy3: String(row[20]),
          confirmedAt3: String(row[21]),
          paidBy: String(row[22]),
          paidAt: String(row[23]),
          notes: String(row[24]),
          files: String(row[25]),
          printed: row[26] === true || row[26] === 'TRUE',
          printedBy: String(row[27] || ''),
          printedAt: String(row[28] || ''),
          archived: row[29] === true || row[29] === 'TRUE',
          comments: parseComments(row[30])
        };
      }
    }
    
    return null;
  } catch (error) {
    Logger.log('❌ Error getting invoice: ' + error);
    return null;
  }
}

// Save new invoice
function saveInvoice(invoiceData) {
  const lock = LockService.getScriptLock();
  
  try {
    lock.waitLock(30000);
    
    Logger.log('💾 Creating new invoice...');
    Logger.log('📋 Invoice data: ' + JSON.stringify(invoiceData));
    
    // Validate input
    if (!invoiceData) {
      throw new Error('Invoice data is undefined');
    }
    
    if (!invoiceData.number) {
      throw new Error('Invoice number is required');
    }
    
    const sheet = getOrCreateSheet();
    const data = sheet.getDataRange().getValues();
    
    // Generate new ID
    let maxId = 0;
    for (let i = 1; i < data.length; i++) {
      const currentId = parseInt(data[i][0]) || 0;
      if (currentId > maxId) {
        maxId = currentId;
      }
    }
    const newId = maxId + 1;
    
    const now = new Date().toLocaleString('ru-RU', {timeZone: 'Asia/Almaty'});
    
    // Prepare row data (31 columns to match headers)
    const rowData = [
      newId,                                      // 1. id
      String(invoiceData.number || ''),           // 2. number
      invoiceData.date || '',                     // 3. date
      String(invoiceData.company || ''),          // 4. company
      String(invoiceData.supplier || ''),         // 5. supplier
      String(invoiceData.supplierBIN || ''),      // 6. supplierBIN
      parseFloat(invoiceData.amount) || 0,        // 7. amount
      String(invoiceData.currency || 'KZT'),      // 8. currency
      String(invoiceData.purpose || ''),          // 9. purpose
      invoiceData.dueDate || '',                  // 10. dueDate
      String(invoiceData.priority || 'Обычный'),  // 11. priority
      'pending',                                  // 12. status
      String(invoiceData.createdBy || ''),        // 13. createdBy
      now,                                        // 14. createdAt
      '',                                         // 15. approvedBy
      '',                                         // 16. approvedAt
      '',                                         // 17. confirmedBy1
      '',                                         // 18. confirmedAt1
      '',                                         // 19. confirmedBy2
      '',                                         // 20. confirmedAt2
      '',                                         // 21. confirmedBy3
      '',                                         // 22. confirmedAt3
      '',                                         // 23. paidBy
      '',                                         // 24. paidAt
      String(invoiceData.notes || ''),            // 25. notes
      String(invoiceData.files || ''),            // 26. files
      false,                                      // 27. printed
      '',                                         // 28. printedBy
      '',                                         // 29. printedAt
      false,                                      // 30. archived
      ''                                          // 31. comments (JSON)
    ];
    
    Logger.log('📝 Row data prepared: ' + rowData.length + ' columns');
    
    // Add row to sheet
    const nextRow = sheet.getLastRow() + 1;
    sheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
    
    // Format the new row
    formatRow(sheet, nextRow, 'pending');
    
    // Update supplier template (если функция существует)
    try {
      if (invoiceData.supplier && typeof updateSupplierTemplate === 'function') {
        updateSupplierTemplate(invoiceData.supplier, invoiceData.supplierBIN);
      }
    } catch (supplierError) {
      Logger.log('⚠️ Supplier update warning: ' + supplierError);
    }
    
    // Log action (если функция существует)
    try {
      if (typeof logAction === 'function') {
        logAction(
          newId,
          { name: invoiceData.createdBy, role: 'Инициатор' },
          'CREATE_INVOICE',
          '',
          'pending',
          'Создан новый счет №' + invoiceData.number
        );
      }
    } catch (logError) {
      Logger.log('⚠️ Log warning: ' + logError);
    }
    
    Logger.log('✅ Invoice created successfully with ID: ' + newId);
    
    return {
      success: true,
      invoiceId: newId,
      message: 'Счет успешно создан'
    };
    
  } catch (error) {
    Logger.log('❌ Error saving invoice: ' + error);
    Logger.log('❌ Error stack: ' + error.stack);
    return {
      success: false,
      error: error.message || 'Не удалось создать счет'
    };
  } finally {
    lock.releaseLock();
  }
}

// Update invoice status (ЕДИНСТВЕННАЯ ПРАВИЛЬНАЯ ВЕРСИЯ)
function updateInvoiceStatus(invoiceId, newStatus, userInfo) {
  const lock = LockService.getScriptLock();
  
  try {
    lock.waitLock(30000);
    
    Logger.log('🔄 Updating invoice ' + invoiceId + ' to ' + newStatus);
    
    const sheet = getOrCreateSheet();
    const data = sheet.getDataRange().getValues();
    
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(invoiceId)) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, error: 'Invoice not found' };
    }
    
    const now = new Date().toLocaleString('ru-RU', {timeZone: 'Asia/Almaty'});
    
    // ОТКЛОНЕНИЕ
    if (newStatus === 'rejected') {
      sheet.getRange(rowIndex, 12).setValue('rejected');
      sheet.getRange(rowIndex, 15).setValue(userInfo.name);
      sheet.getRange(rowIndex, 16).setValue(now);
      
      if (userInfo.rejectionReason) {
        const currentNotes = data[rowIndex-1][24] || '';
        const rejectionNote = '[ОТКЛОНЕН ' + now + ' - ' + userInfo.name + ']: ' + userInfo.rejectionReason;
        const updatedNotes = currentNotes ? currentNotes + '\n\n' + rejectionNote : rejectionNote;
        sheet.getRange(rowIndex, 25).setValue(updatedNotes);
      }
      
      formatRow(sheet, rowIndex, 'rejected');
      
    // ПОДТВЕРЖДЕНИЕ ФИНАНСАМИ (2 человека, любой порядок)
    } else if (newStatus === 'partial_confirmed' || newStatus === 'confirmed') {
      
      // Проверяем: уже подтверждал этот пользователь?
      const alreadyConfirmedBy1 = data[rowIndex-1][16] === userInfo.name;
      const alreadyConfirmedBy2 = data[rowIndex-1][18] === userInfo.name;
      
      if (alreadyConfirmedBy1 || alreadyConfirmedBy2) {
        return { success: false, error: 'Вы уже подтверждали этот счет' };
      }
      
      // Первое подтверждение
      if (!data[rowIndex-1][16]) {
        sheet.getRange(rowIndex, 17).setValue(userInfo.name); // confirmedBy1
        sheet.getRange(rowIndex, 18).setValue(now); // confirmedAt1
        sheet.getRange(rowIndex, 12).setValue('partial_confirmed');
        formatRow(sheet, rowIndex, 'partial_confirmed');
        Logger.log('✅ First confirmation by ' + userInfo.name);
        
      // Второе подтверждение
      } else if (!data[rowIndex-1][18]) {
        sheet.getRange(rowIndex, 19).setValue(userInfo.name); // confirmedBy2
        sheet.getRange(rowIndex, 20).setValue(now); // confirmedAt2
        sheet.getRange(rowIndex, 12).setValue('confirmed');
        formatRow(sheet, rowIndex, 'confirmed');
        Logger.log('✅ Second confirmation by ' + userInfo.name);
        
      } else {
        return { success: false, error: 'Счет уже полностью подтвержден' };
      }
      
    // ОСТАЛЬНЫЕ СТАТУСЫ
    } else {
      sheet.getRange(rowIndex, 12).setValue(newStatus);
      
      switch (newStatus) {
        case 'approved':
          sheet.getRange(rowIndex, 15).setValue(userInfo.name);
          sheet.getRange(rowIndex, 16).setValue(now);
          break;
        case 'paid':
          sheet.getRange(rowIndex, 23).setValue(userInfo.name);
          sheet.getRange(rowIndex, 24).setValue(now);
          break;
      }
      
      formatRow(sheet, rowIndex, newStatus);
    }
    
    Logger.log('✅ Status updated successfully');
    return { success: true };
    
  } catch (error) {
    Logger.log('❌ Error updating status: ' + error);
    return { success: false, error: error.message };
  } finally {
    lock.releaseLock();
  }
}

// Mark invoice as printed
function markInvoicePrinted(invoiceId, userInfo) {
  try {
    const sheet = getOrCreateSheet();
    const data = sheet.getDataRange().getValues();
    
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(invoiceId)) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, error: 'Invoice not found' };
    }
    
    const now = new Date().toLocaleString('ru-RU', {timeZone: 'Asia/Almaty'});
    
    sheet.getRange(rowIndex, 27).setValue(true);           // Column AA (printed)
    sheet.getRange(rowIndex, 28).setValue(userInfo.name);  // Column AB (printedBy)
    sheet.getRange(rowIndex, 29).setValue(now);            // Column AC (printedAt)
    
    return { success: true };
    
  } catch (error) {
    Logger.log('❌ Error marking printed: ' + error);
    return { success: false, error: error.message };
  }
}

// Add comment to invoice
function addComment(invoiceId, userInfo, commentText) {
  try {
    const sheet = getOrCreateSheet();
    const data = sheet.getDataRange().getValues();
    
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(invoiceId)) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, error: 'Invoice not found' };
    }
    
    const currentComments = parseComments(data[rowIndex-1][30]);
    
    const newComment = {
      user: userInfo.name,
      timestamp: new Date().toISOString(),
      text: commentText
    };
    
    currentComments.push(newComment);
    
    sheet.getRange(rowIndex, 31).setValue(JSON.stringify(currentComments)); // Column AE
    
    return { success: true };
    
  } catch (error) {
    Logger.log('❌ Error adding comment: ' + error);
    return { success: false, error: error.message };
  }
}

// ============================================
// HELPER FUNCTIONS
// ============================================

// Get or create sheet
function getOrCreateSheet() {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!sheet) {
      Logger.log('📋 Creating new sheet: ' + CONFIG.SHEET_NAME);
      sheet = spreadsheet.insertSheet(CONFIG.SHEET_NAME);
      createHeaders(sheet);
    }
    
    return sheet;
  } catch (error) {
    Logger.log('❌ Error accessing sheet: ' + error);
    throw new Error('Check SPREADSHEET_ID in CONFIG');
  }
}

// Create headers
function createHeaders(sheet) {
  const headers = [
    'id', 'number', 'date', 'company', 'supplier', 'supplierBIN',
    'amount', 'currency', 'purpose', 'dueDate', 'priority', 'status',
    'createdBy', 'createdAt', 'approvedBy', 'approvedAt',
    'confirmedBy1', 'confirmedAt1', 'confirmedBy2', 'confirmedAt2',
    'confirmedBy3', 'confirmedAt3', 'paidBy', 'paidAt',
    'notes', 'files', 'printed', 'printedBy', 'printedAt',
    'archived', 'comments'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format headers
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
}

// Format row by status
function formatRow(sheet, rowIndex, status) {
  const range = sheet.getRange(rowIndex, 1, 1, 31);
  
  const colors = {
    pending: '#fff3cd',
    approved: '#d1ecf1',
    partial_confirmed: '#fed7aa',
    confirmed: '#d4edda',
    paid: '#e2f4e1',
    rejected: '#f8d7da'
  };
  
  range.setBackground(colors[status] || '#ffffff');
}

// Format date for JavaScript
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
    
    // Return in YYYY-MM-DD format
    return date.getFullYear() + '-' + 
           String(date.getMonth() + 1).padStart(2, '0') + '-' + 
           String(date.getDate()).padStart(2, '0');
  } catch (error) {
    return String(dateValue);
  }
}

// Parse comments JSON
function parseComments(commentsStr) {
  if (!commentsStr || commentsStr === '') return [];
  
  try {
    return JSON.parse(commentsStr);
  } catch (error) {
    return [];
  }
}
