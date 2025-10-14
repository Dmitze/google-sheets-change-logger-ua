function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("📌 Логи змін")
    .addItem("Очистити всі підсвічування", "clearAllHighlights")
    .addItem("Встановити автоочищення (щодня о 23:00)", "installDailyCleanupTrigger") // ← змінено текст
    .addSeparator()
    .addItem("Створити аркуш логів", "setupLogSheet")
    .addItem("Ввести/змінити ім’я користувача", "promptForUsername")
    .addToUi();
}


const LOG_SHEET_NAME = "Лог змін";
const COLOR_ADDED = "#b6d7a8";   // зелений — додано
const COLOR_DELETED = "#ea9999"; // червоний — видалено
const COLOR_CHANGED = "#ffe599"; // жовтий — змінено
const HIGHLIGHT_DURATION_HOURS = 24;

const IGNORED_HEADERS = [
  'Постійний ID',
  'Ідентифікатор',
  'QR-код',
  'QR',
  'link',
  'Посилання на QR'
];


// --- Введення імені користувача ---
function promptForUsername() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt("Введіть своє ім’я або позивний для логів:");
  if (result.getSelectedButton() === ui.Button.OK) {
    const name = result.getResponseText().trim();
    if (name) {
      PropertiesService.getUserProperties().setProperty("username", name);
      ui.alert("Ім’я збережено як: " + name);
    } else {
      ui.alert("Ім’я не може бути порожнім");
    }
  }
}

// --- Головний тригер редагування ---
function onEdit(e) {
  if (!e || typeof e !== 'object' || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() === LOG_SHEET_NAME) return;
  highlightCell(e);
  logCellEdit(e);
}

// --- Підсвічування зміненої клітинки ---
function highlightCell(e) {
  if (!e || !e.range) {
    console.warn("highlightCell: виклик без події — ігноруємо.");
    return;
  }

  const cell = e.range;
  const sheet = cell.getSheet();
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
}

// --- Запис підсвіченої клітинки для подальшого очищення ---
function recordHighlightedCell(sheetName, a1Notation, timestamp) {
  if (!sheetName || !a1Notation || timestamp == null) {
    console.warn("recordHighlightedCell: некоректні дані — ігноруємо.");
    return;
  }
  const key = `highlight_${sheetName}_${a1Notation}`;
  PropertiesService.getScriptProperties().setProperty(key, String(timestamp));
}

// --- Логування зміни ---
function logCellEdit(e) {
  if (!e || !e.range) {
    console.warn("logCellEdit: виклик без події — ігноруємо.");
    return;
  }

  const sheet = e.range.getSheet();
  if (sheet.getName() === LOG_SHEET_NAME) return;

  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) return;

  const col = e.range.getColumn();
  const row = e.range.getRow();
  if (row === 1) return;

  let header = '';
  try {
    header = sheet.getRange(1, col).getValue();
  } catch (err) {
    return;
  }

  const headerNorm = String(header).trim().toLowerCase();
  if (IGNORED_HEADERS.some(name => name.trim().toLowerCase() === headerNorm)) return;

  const username = PropertiesService.getUserProperties().getProperty("username");
  const email = Session.getActiveUser().getEmail();
  const user = username || (email ? email.split('@')[0] : "Анонім");

  const time = new Date();
  const oldValue = e.oldValue != null ? String(e.oldValue) : "";
  const newValue = e.value != null ? String(e.value) : "";

  let changeType = "Змінено";
  if ((oldValue === "" || oldValue === "null") && newValue !== "") {
    changeType = "Додано значення";
  } else if (oldValue !== "" && (newValue === "" || newValue === "null")) {
    changeType = "Видалено значення";
  }

  const cellLink = `=HYPERLINK("#gid=${sheet.getSheetId()}&range=${e.range.getA1Notation()}"; "${e.range.getA1Notation()}")`;

  logSheet.appendRow([
    time,
    user,
    sheet.getName(),
    cellLink,
    changeType,
    oldValue,
    newValue,
    "",
    ""
  ]);
}

// --- Очищення ТІЛЬКИ тимчасового підсвічування (через меню) ---
function clearAllHighlights() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  const allKeys = props.getKeys();
  const highlightKeys = allKeys.filter(key => key.startsWith("highlight_"));

  if (highlightKeys.length === 0) {
    SpreadsheetApp.getUi().alert("Немає підсвічених змін для очищення.");
    return;
  }

  let clearedCount = 0;
  highlightKeys.forEach(key => {
    const parts = key.split('_');
    if (parts.length >= 3) {
      const sheetName = parts[1];
      const a1Notation = parts.slice(2).join('_');
      const sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        try {
          sheet.getRange(a1Notation).setBackground(null);
          clearedCount++;
        } catch (e) {
          // Ігноруємо помилки (клітинка видалена тощо)
        }
      }
    }
    props.deleteProperty(key);
  });

  SpreadsheetApp.getUi().alert(`✅ Очищено ${clearedCount} підсвічених змін.`);
}

// --- Автоматичне очищення старих підсвічувань (старіше 24 годин) ---
function clearOldHighlights() {
  const now = new Date().getTime();
  const cutoff = now - (HIGHLIGHT_DURATION_HOURS * 60 * 60 * 1000);
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const allKeys = props.getKeys();
  const highlightKeys = allKeys.filter(key => key.startsWith("highlight_"));

  highlightKeys.forEach(key => {
    const timestamp = Number(props.getProperty(key));
    if (isNaN(timestamp) || timestamp < cutoff) {
      const parts = key.split('_');
      if (parts.length >= 3) {
        const sheetName = parts[1];
        const a1Notation = parts.slice(2).join('_');
        const sheet = ss.getSheetByName(sheetName);
        if (sheet) {
          try {
            sheet.getRange(a1Notation).setBackground(null);
          } catch (e) {
            // Ігноруємо
          }
        }
      }
      props.deleteProperty(key);
    }
  });
}

// --- Створення аркушу логів ---
function setupLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) {
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
    logSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    logSheet.autoResizeColumns(1, headers[0].length);
  }
}

// --- Встановлення щоденного тригера для очищення ОБОХ підсвічувань о 23:00 за Києвом ---
function installDailyCleanupTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  // Видаляємо старі тригери
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "clearAllHighlightsAtNight") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Створюємо новий тригер на 23:00 за Києвом
  ScriptApp.newTrigger("clearAllHighlightsAtNight")
    .timeBased()
    .inTimezone("Europe/Kiev") // ← КЛЮЧОВОЙ МОМЕНТ!
    .atHour(23)
    .everyDays(1)
    .create();

  SpreadsheetApp.getUi().alert(
    "✅ Тригер автоочищення встановлено!\n" +
    "Щодня о 23:00 за Києвом уся підсвічування буде повністю очищена."
  );
}

// --- Повне очищення ВСІХ підсвічень (викликається щодня о 23:00 за Києвом) ---
function clearAllHighlightsAtNight() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  const allKeys = props.getKeys();
  const highlightKeys = allKeys.filter(key => key.startsWith("highlight_"));

  if (highlightKeys.length === 0) return; // нічого робити

  highlightKeys.forEach(key => {
    const parts = key.split('_');
    if (parts.length >= 3) {
      const sheetName = parts[1];
      const a1Notation = parts.slice(2).join('_');
      const sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        try {
          sheet.getRange(a1Notation).setBackground(null);
        } catch (e) {
          // Ігноруємо помилки (наприклад, якщо аркуш або клітинка видалені)
        }
      }
    }
    props.deleteProperty(key);
  });

  // Опціонально: логування очищення (можна вимкнути)
  console.log(`[Автоочищення] Очищено ${highlightKeys.length} підсвічених клітинок о 23:00 за Києвом.`);
}
