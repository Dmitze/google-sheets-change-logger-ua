const LOG_SHEET_NAME = "Лог змін";
const COLOR_ADDED = "#b6d7a8";
const COLOR_DELETED = "#ea9999";
const COLOR_CHANGED = "#ffe599";
const HIGHLIGHT_DURATION_HOURS = 24;

const IGNORED_HEADERS = [
  'Постійний ID', 'Ідентифікатор', 'QR-код', 'QR', 'link', 'Посилання на QR',
  'ID', 'UUID', 'GUID', 'Timestamp', 'Час створення'
];

// --- Головна функція для обробки змін ---
function onEdit(e) {
  if (!e) {
    console.warn("onEdit: виклик без події");
    return;
  }
  
  try {
    // Виконуємо обидві функції паралельно
    highlightCell(e);
    logCellEdit(e);
  } catch (error) {
    console.error('Помилка в onEdit:', error);
  }
}

// --- Введення імені користувача ---
function promptForUsername() {
  try {
    const ui = SpreadsheetApp.getUi();
    const currentName = PropertiesService.getUserProperties().getProperty("username") || "";
    
    const result = ui.prompt(
      "Введення імені", 
      `Поточне ім'я: ${currentName || "не встановлено"}\n\nВведіть нове ім'я або позивний для логів:`,
      ui.ButtonSet.OK_CANCEL
    );
    
    if (result.getSelectedButton() === ui.Button.OK) {
      const name = result.getResponseText().trim();
      if (name) {
        PropertiesService.getUserProperties().setProperty("username", name);
        ui.alert("✅ Успіх", `Ім'я збережено: ${name}`, ui.ButtonSet.OK);
      } else {
        ui.alert("❌ Помилка", "Ім'я не може бути порожнім", ui.ButtonSet.OK);
      }
    }
  } catch (error) {
    console.error('Помилка при введенні імені:', error);
    SpreadsheetApp.getUi().alert('Помилка при збереженні імені');
  }
}

// --- Підсвічування зміненої клітинки ---
function highlightCell(e) {
  if (!e || !e.range) {
    console.warn("highlightCell: виклик без події");
    return;
  }

  try {
    const cell = e.range;
    const sheet = cell.getSheet();
    
    // Перевіряємо, чи це не аркуш логів
    if (sheet.getName() === LOG_SHEET_NAME) return;

    const oldValue = e.oldValue != null ? String(e.oldValue) : "";
    const newValue = e.value != null ? String(e.value) : "";

    if (oldValue === newValue) return;

    let color;
    if ((oldValue === "" || oldValue === "null") && newValue !== "") {
      color = COLOR_ADDED;
    } else if (oldValue !== "" && (newValue === "" || newValue === "null")) {
      color = COLOR_DELETED;
    } else {
      color = COLOR_CHANGED;
    }

    cell.setBackground(color);
    recordHighlightedCell(sheet.getName(), cell.getA1Notation(), new Date().getTime());
  } catch (error) {
    console.error('Помилка при підсвічуванні:', error);
  }
}

// --- Запис підсвіченої клітинки ---
function recordHighlightedCell(sheetName, a1Notation, timestamp) {
  try {
    if (!sheetName || !a1Notation || timestamp == null) {
      console.warn("recordHighlightedCell: некоректні дані");
      return;
    }
    
    const props = PropertiesService.getScriptProperties();
    const key = `highlight_${sheetName}_${a1Notation.replace(/[\.\[\]]/g, '_')}`;
    
    // Очищаємо старі записи, якщо їх забагато
    const allKeys = props.getKeys().filter(k => k.startsWith("highlight_"));
    if (allKeys.length > 1000) {
      const sortedKeys = allKeys.sort((a, b) => {
        return Number(props.getProperty(b)) - Number(props.getProperty(a));
      });
      
      for (let i = 500; i < sortedKeys.length; i++) {
        props.deleteProperty(sortedKeys[i]);
      }
    }
    
    props.setProperty(key, String(timestamp));
  } catch (error) {
    console.error('Помилка при записі підсвічування:', error);
  }
}

// --- Логування зміни ---
function logCellEdit(e) {
  if (!e || !e.range) {
    console.warn("logCellEdit: виклик без події");
    return;
  }

  try {
    const sheet = e.range.getSheet();
    if (sheet.getName() === LOG_SHEET_NAME) return;

    const logSheet = ensureLogSheet();
    if (!logSheet) {
      console.warn('Аркуш логів не знайдено');
      return;
    }

    const col = e.range.getColumn();
    const row = e.range.getRow();
    
    // Ігноруємо зміни в заголовках
    if (row === 1) return;

    let header = '';
    try {
      header = sheet.getRange(1, col).getValue();
    } catch (err) {
      console.log('Не вдалося отримати заголовок:', err);
      return;
    }

    // Перевіряємо, чи це ігнорований заголовок
    const headerNorm = String(header).trim().toLowerCase();
    if (IGNORED_HEADERS.some(name => name.trim().toLowerCase() === headerNorm)) {
      return;
    }

    const user = getCurrentUserName();
    const time = new Date();
    const oldValue = e.oldValue != null ? String(e.oldValue) : "";
    const newValue = e.value != null ? String(e.value) : "";

    // Визначаємо тип зміни
    let changeType = "Змінено";
    if ((oldValue === "" || oldValue === "null") && newValue !== "") {
      changeType = "Додано значення";
    } else if (oldValue !== "" && (newValue === "" || newValue === "null")) {
      changeType = "Видалено значення";
    }

    // Створюємо посилання на клітинку
    const cellLink = `=HYPERLINK("#gid=${sheet.getSheetId()}&range=${e.range.getA1Notation()}"; "${e.range.getA1Notation()}")`;

    // Додаємо запис в лог
    logSheet.appendRow([
      time,
      user,
      sheet.getName(),
      cellLink,
      changeType,
      oldValue,
      newValue,
      "", // Формула
      ""  // Важлива зміна
    ]);

    // Автоматично налаштовуємо ширину колонок
    logSheet.autoResizeColumns(1, 9);
    
    console.log('Записано зміну в лог:', {
      sheet: sheet.getName(),
      cell: e.range.getA1Notation(),
      user: user,
      type: changeType
    });
    
  } catch (error) {
    console.error('Помилка при логуванні:', error);
  }
}

// --- Отримання імені користувача ---
function getCurrentUserName() {
  try {
    const username = PropertiesService.getUserProperties().getProperty("username");
    if (username) return username;
    
    const email = Session.getEffectiveUser().getEmail();
    return email ? email.split('@')[0] : "Анонім";
  } catch (error) {
    return "Анонім";
  }
}

// --- Створення/перевірка аркуша логів ---
function setupLogSheet() {
  try {
    const logSheet = ensureLogSheet(true);
    if (logSheet) {
      SpreadsheetApp.getUi().alert("✅ Успіх", "Аркуш логів готовий до використання!", SpreadsheetApp.getUi().ButtonSet.OK);
    }
  } catch (error) {
    console.error('Помилка при створенні аркуша логів:', error);
    SpreadsheetApp.getUi().alert('❌ Помилка', 'Не вдалося створити аркуш логів. Перевірте права доступу.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function ensureLogSheet(forceSetup = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  
  if (!logSheet || forceSetup) {
    try {
      if (logSheet) {
        ss.deleteSheet(logSheet);
      }
      
      logSheet = ss.insertSheet(LOG_SHEET_NAME);
      
      const headers = [[
        "Час зміни",
        "Користувач",
        "Аркуш",
        "Посилання на комірку",
        "Тип дії",
        "Було",
        "Стало",
        "Формула",
        "Важлива зміна"
      ]];
      
      const headerRange = logSheet.getRange(1, 1, 1, headers[0].length);
      headerRange.setValues(headers);
      headerRange.setBackground("#dddddd")
                .setFontWeight("bold")
                .setHorizontalAlignment("center");
      
      logSheet.setFrozenRows(1);
      logSheet.autoResizeColumns(1, headers[0].length);
      
      const filterRange = logSheet.getRange(1, 1, 1, headers[0].length);
      filterRange.createFilter();
      
      try {
        const protection = logSheet.protect();
        protection.setWarningOnly(true);
        protection.setDescription("Захист логів - редагування дозволено");
      } catch (protectError) {
        console.log('Не вдалося налаштувати захист аркуша:', protectError);
      }
      
    } catch (error) {
      console.error('Помилка при створенні аркуша логів:', error);
      return null;
    }
  }
  
  return logSheet;
}

// --- Очищення підсвічувань ---
function clearAllHighlights() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const props = PropertiesService.getScriptProperties();
    const allKeys = props.getKeys();
    const highlightKeys = allKeys.filter(key => key.startsWith("highlight_"));

    if (highlightKeys.length === 0) {
      SpreadsheetApp.getUi().alert("ℹ️ Інформація", "Немає підсвічених змін для очищення.", SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    let clearedCount = 0;
    highlightKeys.forEach(key => {
      try {
        const parts = key.split('_');
        if (parts.length >= 3) {
          const sheetName = parts[1];
          const a1Notation = parts.slice(2).join('_').replace(/_/g, '.');
          const sheet = ss.getSheetByName(sheetName);
          if (sheet) {
            sheet.getRange(a1Notation).setBackground(null);
            clearedCount++;
          }
        }
        props.deleteProperty(key);
      } catch (e) {
        // Продовжуємо видаляти наступні клітинки
      }
    });

    SpreadsheetApp.getUi().alert("✅ Успіх", `Очищено ${clearedCount} підсвічених змін.`, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    console.error('Помилка при очищенні підсвічувань:', error);
    SpreadsheetApp.getUi().alert('❌ Помилка', 'Не вдалося очистити підсвічування.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// --- Автоматичне очищення старих підсвічувань ---
function clearOldHighlights() {
  try {
    const now = new Date().getTime();
    const cutoff = now - (HIGHLIGHT_DURATION_HOURS * 60 * 60 * 1000);
    const props = PropertiesService.getScriptProperties();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const allKeys = props.getKeys();
    const highlightKeys = allKeys.filter(key => key.startsWith("highlight_"));

    let clearedCount = 0;
    highlightKeys.forEach(key => {
      const timestamp = Number(props.getProperty(key));
      if (isNaN(timestamp) || timestamp < cutoff) {
        try {
          const parts = key.split('_');
          if (parts.length >= 3) {
            const sheetName = parts[1];
            const a1Notation = parts.slice(2).join('_').replace(/_/g, '.');
            const sheet = ss.getSheetByName(sheetName);
            if (sheet) {
              sheet.getRange(a1Notation).setBackground(null);
              clearedCount++;
            }
          }
          props.deleteProperty(key);
        } catch (e) {
          // Продовжуємо з наступними
        }
      }
    });
    
    console.log(`Автоочищення: видалено ${clearedCount} підсвічувань`);
  } catch (error) {
    console.error('Помилка при автоочищенні:', error);
  }
}

// --- Керування тригерами ---
function installDailyCleanupTrigger() {
  try {
    removeDailyCleanupTrigger();
    
    ScriptApp.newTrigger("clearOldHighlights")
      .timeBased()
      .atHour(23)
      .everyDays(1)
      .create();

    SpreadsheetApp.getUi().alert("✅ Успіх", "Тригер автоочищення встановлено!\nЩодня о 23:00 (за Києвом) підсвічування буде очищено.", SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    console.error('Помилка при встановленні тригера:', error);
    SpreadsheetApp.getUi().alert('❌ Помилка', 'Не вдалося встановити тригер автоочищення.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function removeDailyCleanupTrigger() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let removedCount = 0;
    
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === "clearOldHighlights") {
        ScriptApp.deleteTrigger(trigger);
        removedCount++;
      }
    });
    
    if (removedCount > 0) {
      SpreadsheetApp.getUi().alert("✅ Успіх", `Видалено ${removedCount} тригерів автоочищення.`, SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      SpreadsheetApp.getUi().alert("ℹ️ Інформація", "Тригери автоочищення не знайдено.", SpreadsheetApp.getUi().ButtonSet.OK);
    }
  } catch (error) {
    console.error('Помилка при видаленні тригера:', error);
    SpreadsheetApp.getUi().alert('❌ Помилка', 'Не вдалося видалити тригер автоочищення.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// --- Додаткові утиліти ---
function showLogs() {
  try {
    const logSheet = ensureLogSheet();
    if (logSheet) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      ss.setActiveSheet(logSheet);
    }
  } catch (error) {
    console.error('Помилка при відкритті логів:', error);
  }
}

function testLogging() {
  try {
    const ui = SpreadsheetApp.getUi();
    const result = ui.alert(
      "Тестування логування",
      "Це створить тестову зміну в поточній клітинці для перевірки роботи логів. Продовжити?",
      ui.ButtonSet.YES_NO
    );
    
    if (result === ui.YES) {
      const sheet = SpreadsheetApp.getActiveSheet();
      const range = sheet.getActiveCell();
      const oldValue = range.getValue();
      
      // Створюємо тестову зміну
      range.setValue(oldValue + " [тест]");
      
      ui.alert("✅ Тест завершено", "Перевірте аркуш логів для підтвердження роботи.", ui.ButtonSet.OK);
    }
  } catch (error) {
    console.error('Помилка при тестуванні:', error);
    SpreadsheetApp.getUi().alert('❌ Помилка', 'Не вдалося виконати тест.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
