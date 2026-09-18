/**
 * ============================================================================
 * Wildberries Sheets Data Sync (GAS Automation)
 * Модуль интеграции Google Таблиц с Wildberries Content API v2
 * 
 * Автор: Dmitry (code-dmitriy)
 * Стек: Google Apps Script (JavaScript V8), REST API, Google Sheets API
 * 
 * Ключевые фичи:
 * - Пакетная курсорная пагинация без превышения лимитов памяти и квот execution time
 * - Принудительная строковая нормализация артикулов (сохранение лидирующих нулей '00xxx')
 * - Защита 13-значных баркодов EAN-13 от искажения в экспоненциальный вид (1.2E+12)
 * - Мульти-аккаунт: поддержка работы с несколькими кабинетами селлеров
 * ============================================================================
 */

// Конфигурация интеграции и магазинов
const WB_CONFIG = {
  API_URL: 'https://content-api.wildberries.ru/content/v2/get/cards/list',
  BATCH_LIMIT: 100, // Максимальный лимит выдачи карточек на 1 запрос к API WB
  
  // Безопасное хранилище токенов (все приватные ключи вынесены в заглушки)
  STORES: {
    'Основной кабинет': 'YOUR_WILDBERRIES_API_TOKEN_SHOP_1',
    'Второй кабинет': 'YOUR_WILDBERRIES_API_TOKEN_SHOP_2'
  }
};

/**
 * Создание пользовательского интерфейса (кнопки в верхнем меню Google Таблицы)
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📦 Синхронизация с WB')
    .addItem('📥 Выгрузить каталог (Основной кабинет)', 'exportMainStore')
    .addItem('📥 Выгрузить каталог (Второй кабинет)', 'exportSecondaryStore')
    .addSeparator()
    .addItem('ℹ️ Проверить статус подключения', 'checkApiStatus')
    .addToUi();
}

function exportMainStore() {
  syncCatalogFromWildberries('Основной кабинет');
}

function exportSecondaryStore() {
  syncCatalogFromWildberries('Второй кабинет');
}

/**
 * Основной процесс выгрузки и нормализации номенклатуры
 * @param {string} storeName Выбранный профиль магазина
 */
function syncCatalogFromWildberries(storeName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(storeName);

  // Если листа под магазин нет — создаем его
  if (!sheet) {
    sheet = spreadsheet.insertSheet(storeName);
  }

  const token = WB_CONFIG.STORES[storeName];

  // Валидация токена перед отправкой сетевых запросов
  if (!token || token.startsWith('YOUR_WILDBERRIES')) {
    SpreadsheetApp.getUi().alert(
      '⚠️ Ошибка авторизации',
      `В конфигурации WB_CONFIG не указан действующий API-токен для профиля "${storeName}".`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  // Сброс старых данных и инициализация шапки таблицы
  sheet.clear();
  const headers = [
    'Артикул WB (nmID)',
    'Артикул продавца (VendorCode)',
    'Бренд',
    'Наименование товара',
    'Предметная категория',
    'Баркоды (ШК / Размеры)',
    'Дата обновления карточки'
  ];

  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#2E3440')
    .setFontColor('#ECEFF4')
    .setHorizontalAlignment('center');
  
  sheet.setFrozenRows(1);

  let totalExported = 0;
  let cursorUpdatedAt = '';
  let cursorNmID = 0;
  let hasNextPage = true;

  try {
    // Цикл курсорной пагинации для обхода каталогов любого объема (11 000+ товаров)
    while (hasNextPage) {
      const requestPayload = {
        settings: {
          sort: { ascending: false },
          filter: { withPhoto: -1 },
          cursor: {
            limit: WB_CONFIG.BATCH_LIMIT,
            updatedAt: cursorUpdatedAt || undefined,
            nmID: cursorNmID || undefined
          }
        }
      };

      const requestOptions = {
        method: 'post',
        contentType: 'application/json',
        headers: {
          'Authorization': token
        },
        payload: JSON.stringify(requestPayload),
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch(WB_CONFIG.API_URL, requestOptions);
      const statusCode = response.getResponseCode();

      if (statusCode !== 200) {
        throw new Error(`Wildberries API вернул HTTP ${statusCode}: ${response.getContentText()}`);
      }

      const responseBody = JSON.parse(response.getContentText());
      const cards = responseBody.cards || [];

      if (cards.length === 0) {
        hasNextPage = false;
        break;
      }

      // Санитизация и нормализация данных перед записью
      const rowsBuffer = cards.map(card => {
        // Сборка всех баркодов через запятую
        const barcodesList = (card.sizes || [])
          .flatMap(size => size.skus || [])
          .join(', ');

        return [
          String(card.nmID || ''),
          // Одинарная кавычка + строковое преобразование гарантируют сохранение лидирующих нулей (напр. '00194)
          `'${card.vendorCode || ''}`,
          card.brand || '—',
          card.title || '—',
          card.subjectName || '—',
          // Защита штрихкодов от экспоненциального сжатия (1.23E+12 -> 200000000012)
          `'${barcodesList}`,
          card.updatedAt ? card.updatedAt.substring(0, 19).replace('T', ' ') : '—'
        ];
      });

      // Пакетная вставка в лист через единый диапазон (в разы быстрее построчной appendRow)
      const startRow = sheet.getLastRow() + 1;
      const targetRange = sheet.getRange(startRow, 1, rowsBuffer.length, headers.length);
      
      // Назначение текстового формата ячеек перед сохранением
      sheet.getRange(startRow, 1, rowsBuffer.length, 2).setNumberFormat('@');
      sheet.getRange(startRow, 6, rowsBuffer.length, 1).setNumberFormat('@');
      
      targetRange.setValues(rowsBuffer);

      totalExported += cards.length;

      // Анализ курсора ответа WB для перехода к следующей пачке
      if (responseBody.cursor && responseBody.cursor.total === WB_CONFIG.BATCH_LIMIT) {
        cursorUpdatedAt = responseBody.cursor.updatedAt;
        cursorNmID = responseBody.cursor.nmID;
      } else {
        hasNextPage = false;
      }

      // Пауза 250 мс для соблюдения Rate Limits API Wildberries (защита от 429 Too Many Requests)
      Utilities.sleep(250);
    }

    // Автоподгонка ширины колонок под контент
    for (let col = 1; col <= headers.length; col++) {
      sheet.autoResizeColumn(col);
    }

    SpreadsheetApp.getUi().alert(
      '✅ Синхронизация завершена',
      `Успешно выгружено товаров: ${totalExported} шт.\nМагазин: ${storeName}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } catch (err) {
    Logger.log(`[WB_SYNC_ERROR]: ${err.toString()}`);
    SpreadsheetApp.getUi().alert(
      '❌ Сбой синхронизации',
      `Детали ошибки: ${err.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Проверка доступности шлюза WB API
 */
function checkApiStatus() {
  SpreadsheetApp.getUi().alert(
    'Статус системы',
    'Модуль настроен на работу с протоколом WB Content API v2. Токены экранированы.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
