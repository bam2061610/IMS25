// ============================================
// Constants.gs - Константы системы
// ============================================

// Индексы колонок в таблице (начиная с 0)
const COLUMN = {
  ID: 0,
  NUMBER: 1,
  DATE: 2,
  COMPANY: 3,
  SUPPLIER: 4,
  SUPPLIER_BIN: 5,
  AMOUNT: 6,
  CURRENCY: 7,
  PURPOSE: 8,
  DUE_DATE: 9,
  PRIORITY: 10,
  STATUS: 11,
  CREATED_BY: 12,
  CREATED_AT: 13,
  APPROVED_BY: 14,
  APPROVED_AT: 15,
  CONFIRMED_BY_1: 16,
  CONFIRMED_AT_1: 17,
  CONFIRMED_BY_2: 18,
  CONFIRMED_AT_2: 19,
  CONFIRMED_BY_3: 20,
  CONFIRMED_AT_3: 21,
  PAID_BY: 22,
  PAID_AT: 23,
  NOTES: 24,
  FILES: 25,
  PRINTED: 26,
  PRINTED_BY: 27,
  PRINTED_AT: 28,
  ARCHIVED: 29,
  COMMENTS: 30
};

// Статусы счетов
const STATUS = {
  PENDING: 'pending',
  APPROVED: 'approved',
  PARTIAL_CONFIRMED: 'partial_confirmed',
  CONFIRMED: 'confirmed',
  PAID: 'paid',
  REJECTED: 'rejected'
};

// Цвета для статусов
const STATUS_COLORS = {
  pending: '#fff3cd',
  approved: '#d1ecf1',
  partial_confirmed: '#fed7aa',
  confirmed: '#d4edda',
  paid: '#e2f4e1',
  rejected: '#f8d7da'
};

// Названия статусов
const STATUS_LABELS = {
  pending: 'На согласовании',
  approved: 'Согласован',
  partial_confirmed: 'Частично подтвержден',
  confirmed: 'Подтвержден к оплате',
  paid: 'Оплачен',
  rejected: 'Отклонен'
};

// Приоритеты
const PRIORITY = {
  NORMAL: 'Обычный',
  HIGH: 'Высокий',
  URGENT: 'Срочно'
};

// Валюты
const CURRENCY = {
  KZT: 'KZT',
  USD: 'USD',
  EUR: 'EUR',
  RUB: 'RUB'
};

// Лимиты
const LIMITS = {
  MAX_FILE_SIZE: 20 * 1024 * 1024, // 20 MB
  MAX_FILES_COUNT: 3,
  MAX_COMMENT_LENGTH: 1000,
  MAX_NOTES_LENGTH: 5000
};

// Таймауты и интервалы
const TIMEOUTS = {
  LOCK_WAIT: 30000, // 30 секунд
  CACHE_DURATION: 300, // 5 минут
  RETRY_DELAY: 1000 // 1 секунда
};

// Rate limits
const RATE_LIMITS = {
  CREATE_INVOICE: { calls: 10, period: 60000 }, // 10 в минуту
  UPDATE_STATUS: { calls: 20, period: 60000 },
  UPLOAD_FILE: { calls: 5, period: 60000 }
};

// Сообщения об ошибках
const ERROR_MESSAGES = {
  INVOICE_NOT_FOUND: 'Счет не найден',
  INVALID_DATA: 'Некорректные данные',
  PERMISSION_DENIED: 'Недостаточно прав',
  ALREADY_CONFIRMED: 'Вы уже подтверждали этот счет',
  FILE_TOO_LARGE: 'Файл слишком большой',
  TOO_MANY_FILES: 'Слишком много файлов',
  RATE_LIMIT_EXCEEDED: 'Превышен лимит запросов'
};

// Роли пользователей
const PERMISSIONS = {
  CREATE: 'create',
  APPROVE: 'approve',
  CONFIRM: 'confirm',
  PAY: 'pay',
  VIEW: 'view',
  ALL: 'all'
};

// Типы действий для логирования
const ACTION_TYPES = {
  CREATE_INVOICE: 'CREATE_INVOICE',
  UPDATE_STATUS: 'UPDATE_STATUS',
  ADD_COMMENT: 'ADD_COMMENT',
  UPLOAD_FILE: 'UPLOAD_FILE',
  MARK_PRINTED: 'MARK_PRINTED',
  ARCHIVE: 'ARCHIVE'
};

// Настройки таймзоны
const TIMEZONE = 'Asia/Almaty';

// Форматы дат
const DATE_FORMATS = {
  DATE_ONLY: 'YYYY-MM-DD',
  DATETIME: 'YYYY-MM-DD HH:mm:ss',
  DISPLAY: 'DD.MM.YYYY'
};
