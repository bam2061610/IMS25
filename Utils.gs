// ============================================
// Utils.gs - Утилиты и вспомогательные функции
// ============================================

/**
 * Валидация данных счета
 */
function validateInvoiceData(data) {
  const errors = [];
  
  if (!data) {
    errors.push('Данные счета отсутствуют');
    return errors;
  }
  
  // Обязательные поля
  if (!data.number || String(data.number).trim() === '') {
    errors.push('Номер счета обязателен');
  }
  
  if (!data.company || String(data.company).trim() === '') {
    errors.push('Компания обязательна');
  }
  
  if (!data.supplier || String(data.supplier).trim() === '') {
    errors.push('Поставщик обязателен');
  }
  
  if (!data.amount || isNaN(parseFloat(data.amount)) || parseFloat(data.amount) <= 0) {
    errors.push('Некорректная сумма');
  }
  
  if (!data.purpose || String(data.purpose).trim() === '') {
    errors.push('Назначение платежа обязательно');
  }
  
  if (!data.date) {
    errors.push('Дата счета обязательна');
  }
  
  if (!data.dueDate) {
    errors.push('Срок оплаты обязателен');
  }
  
  // Проверка дат
  if (data.date && data.dueDate) {
    const invoiceDate = new Date(data.date);
    const dueDate = new Date(data.dueDate);
    
    if (isNaN(invoiceDate.getTime())) {
      errors.push('Некорректная дата счета');
    }
    
    if (isNaN(dueDate.getTime())) {
      errors.push('Некорректная дата оплаты');
    }
    
    if (invoiceDate > dueDate) {
      errors.push('Срок оплаты не может быть раньше даты счета');
    }
  }
  
  // Проверка длины полей
  if (data.notes && data.notes.length > LIMITS.MAX_NOTES_LENGTH) {
    errors.push(`Примечания слишком длинные (максимум ${LIMITS.MAX_NOTES_LENGTH} символов)`);
  }
  
  if (data.supplierBIN && data.supplierBIN.length !== 12) {
    errors.push('БИН должен содержать 12 цифр');
  }
  
  return errors;
}

/**
 * Валидация файлов
 */
function validateFiles(files) {
  const errors = [];
  
  if (!Array.isArray(files)) {
    errors.push('Некорректный формат файлов');
    return errors;
  }
  
  if (files.length === 0) {
    errors.push('Необходимо прикрепить хотя бы один файл');
    return errors;
  }
  
  if (files.length > LIMITS.MAX_FILES_COUNT) {
    errors.push(`Максимальное количество файлов: ${LIMITS.MAX_FILES_COUNT}`);
  }
  
  files.forEach((file, index) => {
    if (!file.name) {
      errors.push(`Файл ${index + 1}: отсутствует имя`);
    }
    
    if (!file.content) {
      errors.push(`Файл ${index + 1}: отсутствует содержимое`);
    }
    
    if (file.size && file.size > LIMITS.MAX_FILE_SIZE) {
      const sizeMB = (file.size / 1024 / 1024).toFixed(2);
      const maxMB = (LIMITS.MAX_FILE_SIZE / 1024 / 1024).toFixed(2);
      errors.push(`Файл ${file.name}: размер ${sizeMB}MB превышает максимум ${maxMB}MB`);
    }
  });
  
  return errors;
}

/**
 * Получить текущую метку времени
 */
function getCurrentTimestamp() {
  return new Date().toLocaleString('ru-RU', {timeZone: TIMEZONE});
}

/**
 * Безопасное преобразование в число
 */
function safeParseFloat(value, defaultValue = 0) {
  if (value === null || value === undefined || value === '') {
    return defaultValue;
  }
  const parsed = parseFloat(value);
  return isNaN(parsed) ? defaultValue : parsed;
}

/**
 * Безопасное преобразование в целое число
 */
function safeParseInt(value, defaultValue = 0) {
  if (value === null || value === undefined || value === '') {
    return defaultValue;
  }
  const parsed = parseInt(value);
  return isNaN(parsed) ? defaultValue : parsed;
}

/**
 * Безопасное преобразование в строку
 */
function safeString(value, defaultValue = '') {
  if (value === null || value === undefined) {
    return defaultValue;
  }
  return String(value);
}

/**
 * Безопасное преобразование в булево значение
 */
function safeBoolean(value, defaultValue = false) {
  if (value === null || value === undefined || value === '') {
    return defaultValue;
  }
  return value === true || value === 'TRUE' || value === 'true' || value === 1;
}

/**
 * Форматирование даты для JavaScript
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
    
    // Return in YYYY-MM-DD format
    return date.getFullYear() + '-' + 
           String(date.getMonth() + 1).padStart(2, '0') + '-' + 
           String(date.getDate()).padStart(2, '0');
  } catch (error) {
    Logger.log('❌ Error formatting date: ' + error);
    return String(dateValue);
  }
}

/**
 * Парсинг JSON с обработкой ошибок
 */
function safeParseJSON(jsonStr, defaultValue = null) {
  if (!jsonStr || jsonStr === '') {
    return defaultValue;
  }
  
  try {
    return JSON.parse(jsonStr);
  } catch (error) {
    Logger.log('❌ Error parsing JSON: ' + error);
    return defaultValue;
  }
}

/**
 * Конвертация JSON в строку с обработкой ошибок
 */
function safeStringifyJSON(obj, defaultValue = '') {
  if (!obj) {
    return defaultValue;
  }
  
  try {
    return JSON.stringify(obj);
  } catch (error) {
    Logger.log('❌ Error stringifying JSON: ' + error);
    return defaultValue;
  }
}

/**
 * Парсинг комментариев
 */
function parseComments(commentsStr) {
  return safeParseJSON(commentsStr, []);
}

/**
 * Расчет количества дней между датами
 */
function daysBetween(date1, date2) {
  const d1 = new Date(date1);
  const d2 = new Date(date2);
  
  if (isNaN(d1.getTime()) || isNaN(d2.getTime())) {
    return 0;
  }
  
  const diffTime = Math.abs(d2 - d1);
  return Math.floor(diffTime / (1000 * 60 * 60 * 24));
}

/**
 * Проверка просрочки
 */
function isOverdue(dueDateStr) {
  if (!dueDateStr) return false;
  
  try {
    const dueDate = new Date(dueDateStr);
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    dueDate.setHours(0, 0, 0, 0);
    
    return dueDate < today;
  } catch (error) {
    return false;
  }
}

/**
 * Генерация уникального ID
 */
function generateUniqueId() {
  return Utilities.getUuid();
}

/**
 * Обработка ошибки с логированием
 */
function handleError(error, context, additionalInfo = {}) {
  const errorInfo = {
    message: error.message,
    stack: error.stack,
    context: context,
    timestamp: getCurrentTimestamp(),
    ...additionalInfo
  };
  
  Logger.log('❌ Error in ' + context + ':');
  Logger.log(JSON.stringify(errorInfo, null, 2));
  
  // Опционально: отправить уведомление
  // sendErrorNotification(errorInfo);
  
  return {
    success: false,
    error: error.message,
    context: context
  };
}

/**
 * Retry функция с экспоненциальной задержкой
 */
function retryWithBackoff(fn, maxRetries = 3, baseDelay = 1000) {
  let lastError;
  
  for (let i = 0; i < maxRetries; i++) {
    try {
      return fn();
    } catch (error) {
      lastError = error;
      
      if (i < maxRetries - 1) {
        const delay = baseDelay * Math.pow(2, i);
        Logger.log(`⏳ Retry ${i + 1}/${maxRetries} after ${delay}ms`);
        Utilities.sleep(delay);
      }
    }
  }
  
  throw lastError;
}

/**
 * Санитизация строки для безопасного сохранения
 */
function sanitizeString(str) {
  if (!str) return '';
  return String(str).trim().replace(/[<>]/g, '');
}

/**
 * Валидация email
 */
function isValidEmail(email) {
  if (!email) return false;
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

/**
 * Форматирование суммы для отображения
 */
function formatAmount(amount, currency = 'KZT') {
  if (!amount && amount !== 0) return '0.00 ' + currency;
  
  const formatted = parseFloat(amount).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
  return formatted + ' ' + currency;
}

/**
 * Получение значения из кэша с fallback
 */
function getCachedValue(key, fallbackFn, cacheDuration = TIMEOUTS.CACHE_DURATION) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(key);
  
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (error) {
      Logger.log('⚠️ Cache parse error: ' + error);
    }
  }
  
  const value = fallbackFn();
  
  try {
    cache.put(key, JSON.stringify(value), cacheDuration);
  } catch (error) {
    Logger.log('⚠️ Cache put error: ' + error);
  }
  
  return value;
}

/**
 * Инвалидация кэша
 */
function invalidateCache(key) {
  try {
    CacheService.getScriptCache().remove(key);
  } catch (error) {
    Logger.log('⚠️ Cache invalidate error: ' + error);
  }
}

/**
 * Проверка rate limit
 */
function checkRateLimit(action, userId) {
  const cache = CacheService.getUserCache();
  const key = `ratelimit_${action}_${userId}`;
  const count = parseInt(cache.get(key) || '0');
  
  const limit = RATE_LIMITS[action];
  if (!limit) return true;
  
  if (count >= limit.calls) {
    throw new Error(ERROR_MESSAGES.RATE_LIMIT_EXCEEDED);
  }
  
  cache.put(key, String(count + 1), limit.period / 1000);
  return true;
}

/**
 * Batch обновление ячеек
 */
function batchUpdateCells(sheet, updates) {
  if (!updates || updates.length === 0) return;
  
  // updates = [{row, col, value}, ...]
  // Группируем обновления для оптимизации
  const groupedUpdates = {};
  
  updates.forEach(update => {
    const key = `${update.row}`;
    if (!groupedUpdates[key]) {
      groupedUpdates[key] = [];
    }
    groupedUpdates[key].push(update);
  });
  
  // Применяем обновления построчно
  Object.keys(groupedUpdates).forEach(rowKey => {
    const rowUpdates = groupedUpdates[rowKey];
    rowUpdates.forEach(update => {
      sheet.getRange(update.row, update.col).setValue(update.value);
    });
  });
}

/**
 * Экспорт данных в CSV
 */
function exportToCSV(data, filename) {
  if (!data || data.length === 0) {
    throw new Error('Нет данных для экспорта');
  }
  
  const csvContent = data.map(row => 
    row.map(cell => {
      const str = String(cell || '');
      // Экранирование кавычек и добавление кавычек при необходимости
      if (str.includes(',') || str.includes('"') || str.includes('\n')) {
        return '"' + str.replace(/"/g, '""') + '"';
      }
      return str;
    }).join(',')
  ).join('\n');
  
  const blob = Utilities.newBlob(csvContent, 'text/csv', filename);
  return blob;
}
