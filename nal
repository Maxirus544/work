

function importNalFrom02_12ToIzhevskGroups_DigitsOnly_Final_WithLocation() {
  // --- Конфигурация ---
  const SRC_SPREADSHEET_ID = '1Ovm0wFN8Xk4wNRabKjBPbHsAvF5rjQl0yXxwpD3839s';
  const DST_SPREADSHEET_ID = '1v0mxuAW3B3u0lltoAuROaoh489wlcouYg8O5SQTG0bI';
  const SRC_SHEET_NAME = 'НАЛ';
  const DST_SHEET_NAME = 'ЛК/VIP';
  const PROPERTY_KEY = 'NalDataHistory';

  // --- Открытие таблиц и листов ---
  const srcSs = SpreadsheetApp.openById(SRC_SPREADSHEET_ID);
  const dstSs = SpreadsheetApp.openById(DST_SPREADSHEET_ID);
  const srcSheet = srcSs.getSheetByName(SRC_SHEET_NAME);
  const dstSheet = dstSs.getSheetByName(DST_SHEET_NAME);

  if (!srcSheet) throw new Error(`Лист "${SRC_SHEET_NAME}" не найден в таблице-источнике.`);
  if (!dstSheet) throw new Error(`Лист "${DST_SHEET_NAME}" не найден в целевой таблице.`);

  // --- Загрузка или инициализация истории ---
  let historyMap = loadHistory();

  // --- Чтение данных из СОВРЕМЕННОЙ таблицы-источника ---
  const accounts1 = srcSheet.getRange("H11:H80").getValues();
  const sums1     = srcSheet.getRange("I11:I80").getValues();
  const accounts2 = srcSheet.getRange("H86:H203").getValues();
  const sums2     = srcSheet.getRange("I86:I203").getValues();

  // --- Вспомогательные функции ---
  function toNum(x) {
    if (x == null) return 0;
    if (typeof x === 'number') return isFinite(x) ? x : 0;

    const raw = String(x);
    const clean = raw.trim().replace(/\u00A0/g, '').replace(/,/g, '').replace(/[^0-9.\-]+/g, '');
    const n = Number(clean);
    return isFinite(n) ? n : 0;
  }

  function normalizeAccount(acc) {
      if (acc === null || acc === undefined) return null;
      const digitsOnly = String(acc).replace(/[^0-9]/g, '');
      return digitsOnly === '' ? null : digitsOnly;
  }

  function addAmountToHistory(acc, amount, srcKey) {
    const normalizedAcc = normalizeAccount(acc);
    if (normalizedAcc === null) return;

    const val = toNum(amount);
    if (!isFinite(val) || val === 0) return;

    if (!historyMap[normalizedAcc]) {
      historyMap[normalizedAcc] = { src1: [], src2: [] };
    }
    historyMap[normalizedAcc][srcKey].push(val);
  }

  // --- Обновление истории новыми данными ---
  for (let i = 0; i < accounts1.length; i++) {
    addAmountToHistory(accounts1[i][0], sums1[i][0], 'src1');
  }
  for (let j = 0; j < accounts2.length; j++) {
    addAmountToHistory(accounts2[j][0], sums2[j][0], 'src2');
  }

  // --- Сохранение обновленной истории ---
  saveHistory(historyMap);

  // --- Работа с целевой таблицей ---
  const dstLastRow = dstSheet.getLastRow();
  const acctToRow = {}; // Хранит связь: нормализованный_аккаунт (только цифры) -> номер строки

  if (dstLastRow >= 2) {
    const dstData = dstSheet.getRange(2, 6, dstLastRow - 1, 2).getValues(); // Читаем аккаунты из столбца F
    for (let r = 0; r < dstData.length; r++) {
      const acc = dstData[r][0]; // Значение из столбца F
      if (acc) {
        const normalizedAcc = normalizeAccount(acc); // <--- НОРМАЛИЗУЕМ АККАУНТ ИЗ ЦЕЛЕВОЙ ТАБЛИЦЫ
        if (normalizedAcc) {
            acctToRow[normalizedAcc] = r + 2; // Сохраняем связь: нормализованный_аккаунт -> номер строки
        }
      }
    }
  }

  let nextRow = (dstLastRow >= 2) ? dstLastRow + 1 : 2;

  for (const acct in historyMap) { // Итерируемся по нормализованным аккаунтам (только цифры) из истории
    const comp = historyMap[acct];

    let formulaParts = [];
    comp.src1.forEach(sum => { if (sum > 0) formulaParts.push(String(sum)); });
    comp.src2.forEach(sum => { if (sum > 0) formulaParts.push(String(sum)); });

    const formula = formulaParts.length > 0 ? "=" + formulaParts.join("+") : "";

    if (acctToRow[acct] !== undefined) { // Проверяем, есть ли такой нормализованный аккаунт уже в acctToRow
      // Аккаунт найден: обновляем формулу в существующей строке
      const row = acctToRow[acct];
      dstSheet.getRange(row, 8).setFormula(formula); // Столбец H
    } else {
      // Аккаунт не найден: добавляем новую строку
      //  Изменено здесь: Форматируем новый аккаунт 
      dstSheet.getRange(nextRow, 6).setValue("Лк " + acct + " Ижевск"); // Столбец F
      dstSheet.getRange(nextRow, 8).setFormula(formula); // Столбец H
      nextRow++;
    }
  }

  // --- Вспомогательные функции для PropertiesService ---
  function loadHistory() {
    const properties = PropertiesService.getUserProperties();
    const historyString = properties.getProperty(PROPERTY_KEY);
    if (historyString) {
      try {
        const parsedData = JSON.parse(historyString);
        for (const accKey in parsedData) {
            if (parsedData.hasOwnProperty(accKey)) {
                parsedData[accKey].src1 = parsedData[accKey].src1.map(Number);
                parsedData[accKey].src2 = parsedData[accKey].src2.map(Number);
            }
        }
        return parsedData;
      } catch (e) {
        Logger.log("Ошибка парсинга истории: " + e);
        return {};
      }
    } else {
      return {};
    }
  }

  function saveHistory(historyData) {
    const properties = PropertiesService.getUserProperties();
    try {
      properties.setProperty(PROPERTY_KEY, JSON.stringify(historyData));
    } catch (e) {
      Logger.log("Ошибка сохранения истории: " + e);
    }
  }
}














function importNalFrom02_12ToPermGroups_CorrectedColumns() {

 

  // --- Конфигурация ---
  const SRC_SPREADSHEET_ID = '1Ovm0wFN8Xk4wNRabKjBPbHsAvF5rjQl0yXxwpD3839s'; // ID той же таблицы "02.12"
  const DST_SPREADSHEET_ID = '1v0mxuAW3B3u0lltoAuROaoh489wlcouYg8O5SQTG0bI'; // ID целевой таблицы "Группы с наличкой"
  const SRC_SHEET_NAME = 'Пермь юань 3'; // Название листа-источника
  const DST_SHEET_NAME = 'ЛК/VIP'; // Название целевого листа (тот же, что и раньше)
  const PROPERTY_KEY = 'PermDataHistory'; // Уникальный ключ для истории этого листа

  // --- Открытие таблиц и листов ---
  const srcSs = SpreadsheetApp.openById(SRC_SPREADSHEET_ID);
  const dstSs = SpreadsheetApp.openById(DST_SPREADSHEET_ID);
  const srcSheet = srcSs.getSheetByName(SRC_SHEET_NAME);
  const dstSheet = dstSs.getSheetByName(DST_SHEET_NAME);

  if (!srcSheet) throw new Error(`Лист "${SRC_SHEET_NAME}" не найден в таблице-источнике.`);
  if (!dstSheet) throw new Error(`Лист "${DST_SHEET_NAME}" не найден в целевой таблице.`);

  // --- Загрузка или инициализация истории ---
  let historyMap = loadHistory(PROPERTY_KEY); // Передаем ключ для загрузки истории

  // --- Чтение данных с листа "Пермь" ---
  // Диапазоны аккаунтов
  const accounts1 = srcSheet.getRange("B12:B100").getValues();
  const accounts2 = srcSheet.getRange("B107:B195").getValues();
  // Диапазоны сумм
  const sums1     = srcSheet.getRange("C12:C100").getValues();
  const sums2     = srcSheet.getRange("C107:C195").getValues();

  // --- Вспомогательные функции ---
  function toNum(x) {
    Logger.log(historyMap)
    if (x == null) return 0;
    if (typeof x === 'number') return isFinite(x) ? x : 0;

    const raw = String(x);
    const clean = raw.trim().replace(/\u00A0/g, '').replace(/[^0-9.,]+/g, ''); // Удаляем все, кроме цифр, точки и запятой
    const standardized = clean.replace(/,/g, '.'); // Заменяем десятичную запятую на точку
    const n = Number(standardized);
    return isFinite(n) ? n : 0;
  }

  function normalizeAccount(acc) {
      if (acc === null || acc === undefined) return null;
      
      let accStr = String(acc);
      accStr = accStr.replace(/^Лк\s/i, '').trim(); // Удаляем "Лк " и лишние пробелы
      
      const separatorIndex = accStr.search(/[,.]/); // Ищем первую запятую или точку
      
      if (separatorIndex !== -1) {
          const digits = accStr.substring(separatorIndex + 1).replace(/[^0-9]/g, ''); // Берем часть после разделителя и убираем нецифровые
          return digits === '' ? null : digits;
      } else {
          const digits = accStr.replace(/[^0-9]/g, ''); // Если разделителя нет, берем все цифры из оставшейся строки
          return digits === '' ? null : digits;
      }
  }

  function addAmountToHistory(acc, amount, srcKey) {
    const normalizedAcc = normalizeAccount(acc);
    if (normalizedAcc === null) return;

    const val = toNum(amount);
    if (!isFinite(val) || val === 0) return;

    if (!historyMap[normalizedAcc]) {
      historyMap[normalizedAcc] = { src1: [], src2: [] };
    }
    historyMap[normalizedAcc][srcKey].push(val);
  }

  // --- Обновление истории новыми данными ---
  for (let i = 0; i < accounts1.length; i++) {
    addAmountToHistory(accounts1[i][0], sums1[i][0], 'src1');
  }
  for (let j = 0; j < accounts2.length; j++) {
    addAmountToHistory(accounts2[j][0], sums2[j][0], 'src2');
  }

  // --- Сохранение обновленной истории ---
  saveHistory(historyMap, PROPERTY_KEY);

  // --- Работа с целевой таблицей ---
  const dstLastRow = dstSheet.getLastRow();
  const acctToRow = {}; // Хранит связь: нормализованный_аккаунт (только цифры) -> номер строки

  if (dstLastRow >= 2) {
    // --- Читаем аккаунты из СТОЛБЦА A целевой таблицы ---
    const dstData = dstSheet.getRange(2, 1, dstLastRow - 1, 1).getValues(); // Столбец A (1-й столбец)
    for (let r = 0; r < dstData.length; r++) {
      const acc = dstData[r][0]; // Значение из столбца A
      if (acc) {
        const normalizedAcc = normalizeAccount(acc); // <--- НОРМАЛИЗУЕМ АККАУНТ ИЗ ЦЕЛЕВОЙ ТАБЛИЦЫ
        if (normalizedAcc) {
            acctToRow[normalizedAcc] = r + 2; // Сохраняем связь: нормализованный_аккаунт -> номер строки
        }
      }
    }
  }

  let nextRow = (dstLastRow >= 2) ? dstLastRow + 1 : 2;

  for (const acct in historyMap) { // Итерируемся по нормализованным аккаунтам (только цифры) из истории
    const comp = historyMap[acct];

    let formulaParts = [];
    comp.src1.forEach(sum => { if (sum > 0) formulaParts.push(String(sum)); });
    comp.src2.forEach(sum => { if (sum > 0) formulaParts.push(String(sum)); });

    const formula = formulaParts.length > 0 ? "=" + formulaParts.join("+") : "";

    if (acctToRow[acct] !== undefined) { // Проверяем, есть ли такой нормализованный аккаунт уже в acctToRow
      // Аккаунт найден: обновляем формулу в существующей строке
      const row = acctToRow[acct];
      // --- Обновляем формулу в СТОЛБЦЕ C целевой таблицы ---
      dstSheet.getRange(row, 3).setFormula(formula); // Столбец C (3-й столбец)
    } else {
      // Аккаунт не найден: добавляем новую строку
      // --- Записываем новый аккаунт в СТОЛБЕЦ A целевой таблицы ---
      dstSheet.getRange(nextRow, 1).setValue("Лк " + acct + " Ижевск"); // Столбец A (1-й столбец)
      // --- Записываем формулу суммы в СТОЛБЕЦ C целевой таблицы ---
      dstSheet.getRange(nextRow, 3).setFormula(formula); // Столбец C (3-й столбец)
      nextRow++;
    }
  }

  // --- Вспомогательные функции для PropertiesService ---
  function loadHistory(key) {
    const properties = PropertiesService.getUserProperties();
    const historyString = properties.getProperty(key);
    if (historyString) {
      try {
        const parsedData = JSON.parse(historyString);
        for (const accKey in parsedData) {
            if (parsedData.hasOwnProperty(accKey)) {
                parsedData[accKey].src1 = parsedData[accKey].src1.map(Number);
                parsedData[accKey].src2 = parsedData[accKey].src2.map(Number);
            }
        }
        return parsedData;
      } catch (e) {
        Logger.log("Ошибка парсинга истории для ключа '" + key + "': " + e);
        return {};
      }
    } else {
      return {};
    }
  }

  function saveHistory(historyData, key) {
    const properties = PropertiesService.getUserProperties();
    try {
      properties.setProperty(key, JSON.stringify(historyData));
    } catch (e) {
      Logger.log("Ошибка сохранения истории для ключа '" + key + "': " + e);
    }
  }
}








function importNalFrom02_12ToChelnyGroups_DigitsOnly_Final_WithLocation() {
  // --- Конфигурация ---
  // Укажи реальные IDs для таблиц 02.12 (источник) и целевой таблицы "Группы с наличкой"
  const SRC_SPREADSHEET_ID = '1Ovm0wFN8Xk4wNRabKjBPbHsAvF5rjQl0yXxwpD3839s';
  const DST_SPREADSHEET_ID = '1v0mxuAW3B3u0lltoAuROaoh489wlcouYg8O5SQTG0bI';
  const SRC_SHEET_NAME = 'Челны'; // лист в таблице 02.12
  const DST_SHEET_NAME = 'ЛК/VIP'; // целевой лист
  const PROPERTY_KEY = 'ChelnyDataHistory';

  // --- Открытие таблиц и листов ---
  const srcSs = SpreadsheetApp.openById(SRC_SPREADSHEET_ID);
  const dstSs = SpreadsheetApp.openById(DST_SPREADSHEET_ID);
  const srcSheet = srcSs.getSheetByName(SRC_SHEET_NAME);
  const dstSheet = dstSs.getSheetByName(DST_SHEET_NAME);

  if (!srcSheet) throw new Error(`Лист "${SRC_SHEET_NAME}" не найден в таблице-источнике.`);
  if (!dstSheet) throw new Error(`Лист "${DST_SHEET_NAME}" не найден в целевой таблице.`);

  // --- Загрузка или инициализация истории ---
  let historyMap = loadHistory();

  // --- Чтение данных из листа Челны ---
  // Лк: B12:B113 и B120:B227
  // Суммы: C12:C113 и C120:C1227
  const accounts1 = srcSheet.getRange("B12:B113").getValues();
  const sums1     = srcSheet.getRange("C12:C113").getValues();
  const accounts2 = srcSheet.getRange("B120:B227").getValues();
  const sums2     = srcSheet.getRange("C120:C1227").getValues();

  // --- Вспомогательные функции ---
  function toNum(x) {
    if (x == null) return 0;
    if (typeof x === 'number') return isFinite(x) ? x : 0;

    const raw = String(x);
    const clean = raw.trim().replace(/\u00A0/g, '').replace(/,/g, '').replace(/[^0-9.\-]+/g, '');
    const n = Number(clean);
    return isFinite(n) ? n : 0;
  }

  function normalizeAccount(acc) {
      if (acc === null || acc === undefined) return null;
      const digitsOnly = String(acc).replace(/[^0-9]/g, '');
      return digitsOnly === '' ? null : digitsOnly;
  }

  function addAmountToHistory(acc, amount, srcKey) {
    const normalizedAcc = normalizeAccount(acc);
    if (normalizedAcc === null) return;

    const val = toNum(amount);
    if (!isFinite(val) || val === 0) return;

    if (!historyMap[normalizedAcc]) {
      historyMap[normalizedAcc] = { src1: [], src2: [] };
    }
    historyMap[normalizedAcc][srcKey].push(val);
  }

  // --- Обновление истории новыми данными ---
  for (let i = 0; i < accounts1.length; i++) {
    addAmountToHistory(accounts1[i][0], sums1[i][0], 'src1');
  }
  for (let j = 0; j < accounts2.length; j++) {
    addAmountToHistory(accounts2[j][0], sums2[j][0], 'src2');
  }

  // --- Сохранение обновленной истории ---
  saveHistory(historyMap);

  // --- Работа с целевой таблицей ---
  const dstLastRow = dstSheet.getLastRow();
  const acctToRow = {}; // Хранит связь: нормализованный_аккаунт (только цифры) -> номер строки

  if (dstLastRow >= 2) {
    // Читаем существующие Лк из столбца K (11-й столбец)
    const dstData = dstSheet.getRange(2, 11, dstLastRow - 1, 1).getValues();
    for (let r = 0; r < dstData.length; r++) {
      const acc = dstData[r][0]; // Значение из столбца K
      if (acc) {
        const normalizedAcc = normalizeAccount(acc);
        if (normalizedAcc) {
            acctToRow[normalizedAcc] = r + 2; // Сохраняем связь: нормализованный_аккаунт -> номер строки
        }
      }
    }
  }

  let nextRow = (dstLastRow >= 2) ? dstLastRow + 1 : 2;

  for (const acct in historyMap) { // Итерируемся по нормализованным аккаунтам (только цифры) из истории
    const comp = historyMap[acct];

    let formulaParts = [];
    comp.src1.forEach(sum => { if (sum > 0) formulaParts.push(String(sum)); });
    comp.src2.forEach(sum => { if (sum > 0) formulaParts.push(String(sum)); });

    const formula = formulaParts.length > 0 ? "=" + formulaParts.join("+") : "";

    if (acctToRow[acct] !== undefined) { // Аккаунт найден: обновляем формулу в существующей строке
      const row = acctToRow[acct];
      dstSheet.getRange(row, 13).setFormula(formula); // Столбец M
    } else {
      // Аккаунт не найден: добавляем новую строку
      //  Форматируем новый аккаунт в столбец K
      dstSheet.getRange(nextRow, 11).setValue("Лк " + acct + " Челны"); // Столбец K
      dstSheet.getRange(nextRow, 13).setFormula(formula); // Столбец M
      nextRow++;
    }
  }

  // --- Вспомогательные функции для PropertiesService ---
  function loadHistory() {
    const properties = PropertiesService.getUserProperties();
    const historyString = properties.getProperty(PROPERTY_KEY);
    if (historyString) {
      try {
        const parsedData = JSON.parse(historyString);
        for (const accKey in parsedData) {
            if (parsedData.hasOwnProperty(accKey)) {
                parsedData[accKey].src1 = parsedData[accKey].src1.map(Number);
                parsedData[accKey].src2 = parsedData[accKey].src2.map(Number);
            }
        }
        return parsedData;
      } catch (e) {
        Logger.log("Ошибка парсинга истории: " + e);
        return {};
      }
    } else {
      return {};
    }
  }

  function saveHistory(historyData) {
    const properties = PropertiesService.getUserProperties();
    try {
      properties.setProperty(PROPERTY_KEY, JSON.stringify(historyData));
    } catch (e) {
      Logger.log("Ошибка сохранения истории: " + e);
    }
  }
}








function importNalFrom02_12ToUfaGroups_DigitsOnly_Final_WithLocation() {
  // --- Конфигурация ---
  // Укажи реальные IDs для таблиц 02.12 (источник) и целевой таблицы "Группы с наличкой"
  const SRC_SPREADSHEET_ID = '1Ovm0wFN8Xk4wNRabKjBPbHsAvF5rjQl0yXxwpD3839s';
  const DST_SPREADSHEET_ID = '1v0mxuAW3B3u0lltoAuROaoh489wlcouYg8O5SQTG0bI';
  const SRC_SHEET_NAME = 'Уфа 1 Юань'; // лист в таблице 02.12
  const DST_SHEET_NAME = 'ЛК/VIP'; // целевой лист
  const PROPERTY_KEY = 'UfaDataHistory';

  // --- Открытие таблиц и листов ---
  const srcSs = SpreadsheetApp.openById(SRC_SPREADSHEET_ID);
  const dstSs = SpreadsheetApp.openById(DST_SPREADSHEET_ID);
  const srcSheet = srcSs.getSheetByName(SRC_SHEET_NAME);
  const dstSheet = dstSs.getSheetByName(DST_SHEET_NAME);

  if (!srcSheet) throw new Error(`Лист "${SRC_SHEET_NAME}" не найден в таблице-источнике.`);
  if (!dstSheet) throw new Error(`Лист "${DST_SHEET_NAME}" не найден в целевой таблице.`);

  // --- Загрузка или инициализация истории ---
  let historyMap = loadHistory();

  // --- Чтение данных из листа Уфа 1 Юань ---
  // Лк: B12:B100 и B107:B195
  // Суммы: C12:C100 и C107:C195
  const accounts1 = srcSheet.getRange("B12:B100").getValues();
  const sums1     = srcSheet.getRange("C12:C100").getValues();
  const accounts2 = srcSheet.getRange("B107:B195").getValues();
  const sums2     = srcSheet.getRange("C107:C195").getValues();

  // --- Вспомогательные функции ---
  function toNum(x) {
    if (x == null) return 0;
    if (typeof x === 'number') return isFinite(x) ? x : 0;

    const raw = String(x);
    const clean = raw.trim().replace(/\u00A0/g, '').replace(/,/g, '').replace(/[^0-9.\-]+/g, '');
    const n = Number(clean);
    return isFinite(n) ? n : 0;
  }

  function normalizeAccount(acc) {
      if (acc === null || acc === undefined) return null;
      const digitsOnly = String(acc).replace(/[^0-9]/g, '');
      return digitsOnly === '' ? null : digitsOnly;
  }

  function addAmountToHistory(acc, amount, srcKey) {
    const normalizedAcc = normalizeAccount(acc);
    if (normalizedAcc === null) return;

    const val = toNum(amount);
    if (!isFinite(val) || val === 0) return;

    if (!historyMap[normalizedAcc]) {
      historyMap[normalizedAcc] = { src1: [], src2: [] };
    }
    historyMap[normalizedAcc][srcKey].push(val);
  }

  // --- Обновление истории новыми данными ---
  for (let i = 0; i < accounts1.length; i++) {
    addAmountToHistory(accounts1[i][0], sums1[i][0], 'src1');
  }
  for (let j = 0; j < accounts2.length; j++) {
    addAmountToHistory(accounts2[j][0], sums2[j][0], 'src2');
  }

  // --- Сохранение обновленной истории ---
  saveHistory(historyMap);

  // --- Работа с целевой таблицей ---
  const dstLastRow = dstSheet.getLastRow();
  const acctToRow = {}; // Хранит связь: нормализованный_аккаунт (только цифры) -> номер строки

  // Читаем существующие Лк из столбца P (16-й столбец)
  if (dstLastRow >= 2) {
    const dstData = dstSheet.getRange(2, 16, dstLastRow - 1, 1).getValues();
    for (let r = 0; r < dstData.length; r++) {
      const acc = dstData[r][0]; // Значение из столбца P
      if (acc) {
        const normalizedAcc = normalizeAccount(acc);
        if (normalizedAcc) {
            acctToRow[normalizedAcc] = r + 2; // Номер строки
        }
      }
    }
  }

  let nextRow = (dstLastRow >= 2) ? dstLastRow + 1 : 2;

  for (const acct in historyMap) { // Итерируемся по нормализованным аккаунтам (только цифры) из истории
    const comp = historyMap[acct];

    let formulaParts = [];
    comp.src1.forEach(sum => { if (sum > 0) formulaParts.push(String(sum)); });
    comp.src2.forEach(sum => { if (sum > 0) formulaParts.push(String(sum)); });

    const formula = formulaParts.length > 0 ? "=" + formulaParts.join("+") : "";

    if (acctToRow[acct] !== undefined) { // Аккаунт найден: обновляем формулу в существующей строке
      const row = acctToRow[acct];
      dstSheet.getRange(row, 18).setFormula(formula); // Столбец R
    } else {
      // Аккаунт не найден: добавляем новую строку
      //  Форматируем новый аккаунт в столбец P
      dstSheet.getRange(nextRow, 16).setValue("Лк " + acct + " Уфа 1 Юань"); // Столбец P
      dstSheet.getRange(nextRow, 18).setFormula(formula); // Столбец R
      nextRow++;
    }
  }

  // --- Вспомогательные функции для PropertiesService ---
  function loadHistory() {
    const properties = PropertiesService.getUserProperties();
    const historyString = properties.getProperty(PROPERTY_KEY);
    if (historyString) {
      try {
        const parsedData = JSON.parse(historyString);
        for (const accKey in parsedData) {
            if (parsedData.hasOwnProperty(accKey)) {
                parsedData[accKey].src1 = parsedData[accKey].src1.map(Number);
                parsedData[accKey].src2 = parsedData[accKey].src2.map(Number);
            }
        }
        return parsedData;
      } catch (e) {
        Logger.log("Ошибка парсинга истории: " + e);
        return {};
      }
    } else {
      return {};
    }
  }

  function saveHistory(historyData) {
    const properties = PropertiesService.getUserProperties();
    try {
      properties.setProperty(PROPERTY_KEY, JSON.stringify(historyData));
    } catch (e) {
      Logger.log("Ошибка сохранения истории: " + e);
    }
  }
}






function importNalFrom02_12ToKirovGroups_DigitsOnly_Final_WithLocation() {
  // --- Конфигурация ---
  const SRC_SPREADSHEET_ID = '1Ovm0wFN8Xk4wNRabKjBPbHsAvF5rjQl0yXxwpD3839s';
  const DST_SPREADSHEET_ID = '1v0mxuAW3B3u0lltoAuROaoh489wlcouYg8O5SQTG0bI';
  const SRC_SHEET_NAME = 'Киров / Уфа 2'; // лист в таблице 02.12
  const DST_SHEET_NAME = 'ЛК/VIP';        // целевой лист
  const PROPERTY_KEY = 'KirovDataHistory';

  // --- Открытие таблиц и листов ---
  const srcSs = SpreadsheetApp.openById(SRC_SPREADSHEET_ID);
  const dstSs = SpreadsheetApp.openById(DST_SPREADSHEET_ID);
  const srcSheet = srcSs.getSheetByName(SRC_SHEET_NAME);
  const dstSheet = dstSs.getSheetByName(DST_SHEET_NAME);

  if (!srcSheet) throw new Error(`Лист "${SRC_SHEET_NAME}" не найден в таблице-источнике.`);
  if (!dstSheet) throw new Error(`Лист "${DST_SHEET_NAME}" не найден в целевой таблице.`);

  // --- Загрузка или инициализация истории ---
  let historyMap = loadHistory();

  // --- Чтение данных из листа Киров / Уфа 2 ---
  // Лк: B12:B100 и B107:B195
  // Суммы: C12:C100 и C107:C195
  const accounts1 = srcSheet.getRange("B12:B100").getValues();
  const sums1     = srcSheet.getRange("C12:C100").getValues();
  const accounts2 = srcSheet.getRange("B107:B195").getValues();
  const sums2     = srcSheet.getRange("C107:C195").getValues();

  // --- Вспомогательные функции ---
  function toNum(x) {
    if (x == null) return 0;
    if (typeof x === 'number') return isFinite(x) ? x : 0;

    const raw = String(x);
    const clean = raw.trim().replace(/\u00A0/g, '').replace(/,/g, '').replace(/[^0-9.\-]+/g, '');
    const n = Number(clean);
    return isFinite(n) ? n : 0;
  }

  function normalizeAccount(acc) {
      if (acc === null || acc === undefined) return null;
      const digitsOnly = String(acc).replace(/[^0-9]/g, '');
      return digitsOnly === '' ? null : digitsOnly;
  }

  function addAmountToHistory(acc, amount, srcKey) {
    const normalizedAcc = normalizeAccount(acc);
    if (normalizedAcc === null) return;

    const val = toNum(amount);
    if (!isFinite(val) || val === 0) return;

    if (!historyMap[normalizedAcc]) {
      historyMap[normalizedAcc] = { src1: [], src2: [] };
    }
    historyMap[normalizedAcc][srcKey].push(val);
  }

  // --- Обновление истории новыми данными ---
  for (let i = 0; i < accounts1.length; i++) {
    addAmountToHistory(accounts1[i][0], sums1[i][0], 'src1');
  }
  for (let j = 0; j < accounts2.length; j++) {
    addAmountToHistory(accounts2[j][0], sums2[j][0], 'src2');
  }

  // --- Сохранение обновленной истории ---
  saveHistory(historyMap);

  // --- Работа с целевой таблицей ---
  const dstLastRow = dstSheet.getLastRow();
  const acctToRow = {}; // Хранит связь: нормализованный_аккаунт (только цифры) -> номер строки

  // Читаем существующие Лк из столбца U (21-й) в диапазоне 2..N
  if (dstLastRow >= 2) {
    const dstData = dstSheet.getRange(2, 21, dstLastRow - 1, 1).getValues(); // столбец U
    for (let r = 0; r < dstData.length; r++) {
      const acc = dstData[r][0];
      if (acc) {
        const normalizedAcc = normalizeAccount(acc);
        if (normalizedAcc) {
            acctToRow[normalizedAcc] = r + 2; // номер строки
        }
      }
    }
  }

  let nextRow = (dstLastRow >= 2) ? dstLastRow + 1 : 2;

  for (const acct in historyMap) { // Итерируемся по нормализованным аккаунтам (только цифры) из истории
    const comp = historyMap[acct];

    let formulaParts = [];
    comp.src1.forEach(sum => { if (sum > 0) formulaParts.push(String(sum)); });
    comp.src2.forEach(sum => { if (sum > 0) formulaParts.push(String(sum)); });

    const formula = formulaParts.length > 0 ? "=" + formulaParts.join("+") : "";

    if (acctToRow[acct] !== undefined) { // Аккаунт найден: обновляем формулу в существующей строке
      const row = acctToRow[acct];
      dstSheet.getRange(row, 23).setFormula(formula); // столбец W
    } else {
      // Аккаунт не найден: добавляем новую строку
      // Форматируем новый аккаунт в столбец U
      dstSheet.getRange(nextRow, 21).setValue("Лк " + acct + " Киров / Уфа 2"); // столбец U
      dstSheet.getRange(nextRow, 23).setFormula(formula); // столбец W
      nextRow++;
    }
  }

  // --- Вспомогательные функции для PropertiesService ---
  function loadHistory() {
    const properties = PropertiesService.getUserProperties();
    const historyString = properties.getProperty(PROPERTY_KEY);
    if (historyString) {
      try {
        const parsedData = JSON.parse(historyString);
        for (const accKey in parsedData) {
            if (parsedData.hasOwnProperty(accKey)) {
                parsedData[accKey].src1 = parsedData[accKey].src1.map(Number);
                parsedData[accKey].src2 = parsedData[accKey].src2.map(Number);
            }
        }
        return parsedData;
      } catch (e) {
        Logger.log("Ошибка парсинга истории: " + e);
        return {};
      }
    } else {
      return {};
    }
  }

  function saveHistory(historyData) {
    const properties = PropertiesService.getUserProperties();
    try {
      properties.setProperty(PROPERTY_KEY, JSON.stringify(historyData));
    } catch (e) {
      Logger.log("Ошибка сохранения истории: " + e);
    }
  }
}









 function clearHistory() {
      PropertiesService.getUserProperties().deleteProperty('NalDataHistory'); // Или .deleteProperty('nalDataHistory') для getScriptProperties()
      Logger.log("История успешно очищена.");
}

    function clearPermHistory() {
  PropertiesService.getUserProperties().deleteProperty('PermDataHistory'); // Удаляем историю, связанную с листом "Пермь"
  Logger.log("История для листа 'Пермь' успешно очищена.");
}

function clearChelnyHistory() {
  PropertiesService.getUserProperties().deleteProperty('ChelnyDataHistory'); // Удаляем историю, связанную с листом "Челны"
  Logger.log("История для листа 'Челны' успешно очищена.");
}

function clearUfaHistory() {
  PropertiesService.getUserProperties().deleteProperty('UfaDataHistory'); // Удаляем историю, связанную с листом "Уфа"
  Logger.log("История для листа 'Уфа' успешно очищена.");
}

function clearKirovHistory() {
  PropertiesService.getUserProperties().deleteProperty('KirovDataHistory'); // Удаляем историю, связанную с листом "Киров"
  Logger.log("История для листа 'Киров' успешно очищена.");
}
