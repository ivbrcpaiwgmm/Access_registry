/**
 * Объект с данными доступа.
 */
const ACCESS_DATA = {
  CREATOR: ["Создатель", "#00ff00"],        // зеленый
  EDITOR: ["Редактор", "#00ff00"],          // зеленый
  COMMENTATOR: ["Комментатор", "#ffff00"],  // желтый
  READER: ["Читатель", "#ff0000"]           // красный
};

/**
 * Главная функция, запускающая основной скрипт. Запускает создание UI. Запускает триггер.
 */
function run() {
  main();
  onOpen();
  createTrigger();
}

/**
 * Основная функция для обработки данных.
 * @throws {Error} Если произошла ошибка во время выполнения.
 */
function main() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const lastRow = sheet.getLastRow();
    const rangeToClear = sheet.getRange(2, 2, lastRow, sheet.getLastColumn());
    rangeToClear.clear({ formatOnly: true, contentsOnly: true, commentsOnly: true });
    const driveLinks = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

    for (let index = 0; index < driveLinks.length; index++) {
      const link = driveLinks[index][0];

      if (isGoogleDocLink(link)) {
        try {
          processDriveLink(sheet, link, index + 2);
        } catch (processError) {
          sheet.getRange(index + 2, 2).setValue("Нет доступа");
        }
      } else {
        sheet.getRange(index + 2, 2).setValue("Некорректная ссылка");
      }
    }
  } catch (error) {
    handleUpdateError(error);
  }
}

/**
 * Функция для обработки данных документа по ссылке.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Таблица, в которой происходит обработка.
 * @param {string} link - Ссылка на документ.
 * @param {number} rowIndex - Номер строки для записи данных.
 * @throws {Error} Если произошла ошибка во время выполнения.
 */
function processDriveLink(sheet, link, rowIndex) {
  const itemId = getItemIdFromUrl(link);

  const usersData = getUsersData(itemId);
  for (let i = 0; i < usersData.length; i++) {
    const columnIndex = i + 2;
    const cell = sheet.getRange(rowIndex, columnIndex);
    cell.setValue(usersData[i].email !== undefined ? usersData[i].email : "Общий доступ");
    const [role, color] = ACCESS_DATA[usersData[i].accessLevel] || [undefined, undefined];
    applyCellStyles(cell, color, role);
  }
}

/**
 * Функция для применения стилей ячейки.
 * @param {GoogleAppsScript.Spreadsheet.Range} cell - Ячейка для применения стилей.
 * @param {string} color - Цвет фона ячейки.
 * @param {string} role - Роль пользователя.
 */
function applyCellStyles(cell, color, role) {
  if (cell.getValue() === "Общий доступ" && color === "#ff0000") {
    cell.setBackground("#FFA500"); // Оранжевый цвет
    cell.setNote("Уровень доступа: Читатель / Комментатор");
  } else {
    cell.setBackground(color);
    cell.setNote(`Уровень доступа: ${role}`);
  }
}

/**
 * Функция для получения данных пользователей по ID документа.
 * @param {string} itemId - ID документа.
 * @returns {AccessData[]} Массив с данными пользователей.
 */
function getUsersData(itemId) {
  const usersData = new Set();
  const permissions = Drive.Permissions.list(itemId).items;

  const creatorEmail = permissions.find(permission => permission.role === 'owner').emailAddress;
  usersData.add({ email: creatorEmail, accessLevel: "CREATOR" });

  permissions.filter(permission => permission.role === 'writer' || permission.role === 'organizer')
    .forEach(editor => usersData.add({ email: editor.emailAddress, accessLevel: "EDITOR" }));

  getCommentators(itemId).forEach(commenter => usersData.add({ email: commenter, accessLevel: "COMMENTATOR" }));

  permissions.filter(permission => permission.role === 'reader')
    .forEach(reader => {
      const email = reader.emailAddress;
      if (!getCommentators(itemId).includes(email)) {
        usersData.add({ email: email, accessLevel: "READER" });
      }
    });
  return Array.from(usersData);
}

/**
 * Функция для проверки, является ли ссылка на документ Google Docs.
 * @param {string} link - Ссылка на документ.
 * @returns {boolean} true, если ссылка ведет на Google Docs, в противном случае - false.
 */
function isGoogleDocLink(link) {
  const match = /\/(?:folders|d)\/([^\/?]+)/.exec(link);
  return !!match;
}

/**
 * Функция для получения списка комментаторов по ID документа.
 * @param {string} itemId - ID документа.
 * @returns {string[]} Массив с электронными адресами комментаторов.
 */
function getCommentators(itemId) {
  const file = DriveApp.getFileById(itemId);
  const commentatorsList = file.getViewers()
    .filter(user => file.getAccess(user) === DriveApp.Permission.COMMENT);
  return commentatorsList.map(commenter => commenter.getEmail());
}

/**
 * Функция для извлечения ID документа из URL.
 * @param {string} url - URL документа.
 * @returns {string|null} ID документа или null, если ID не найден.
 */
function getItemIdFromUrl(url) {
  const match = /\/(?:folders|d)\/([^\/?]+)/.exec(url);
  return match && match[1];
}

/**
 * Обработчик ошибок.
 * @param {Error} error - Объект ошибки.
 */
function handleUpdateError(error) {
  Logger.log("Произошла ошибка: " + error.message);
  // Здесь можно добавить логику обработки ошибок
}

/**
 * Функция для создания пользовательского меню в интерфейсе таблицы.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Мои скрипты')
    .addItem('Обновить Реестр доступов', 'main')
    .addToUi();
}

/**
 * Функция для создания триггера на выполнение основной функции каждый час.
 */
function createTrigger() {
  ScriptApp.newTrigger('main')
    .timeBased()
    .everyHours(1)
    .create();
}
