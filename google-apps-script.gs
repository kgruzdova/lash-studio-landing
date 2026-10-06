/**
 * ВАЖНО: backend для формы онлайн-записи лендинга.
 *
 * КАК ПОДКЛЮЧИТЬ:
 * 1. Откройте https://script.google.com и создайте новый проект
 *    (Extensions -> Apps Script / Расширения -> Apps Script).
 * 2. Удалите содержимое файла Code.gs и вставьте сюда этот код.
 * 3. Нажмите «Развернуть» (Deploy) -> «Новое развертывание» ->
 *    тип «Веб-приложение» (Web app).
 *    - Запуск от имени: «Я» (вы).
 *    - Кто имеет доступ: «Все, у кого есть ссылка» (Anyone).
 * 4. Скопируйте URL веб-приложения и вставьте его в лендинг:
 *    в index.html в блоке «ВАЖНО: backend пока не подключён...»
 *    замените обработчик submit на вызов fetch по этому URL
 *    (пример кода приведён в функции sendToBackend ниже).
 *
 * ФОРМАТ ЗАПИСИ В ТАБЛИЦЕ:
 * A1 — дата заявки (автоматически)
 * B  — имя
 * C  — телефон
 * D  — эффект
 * E  — желаемая дата
 * F  — удобное время
 * G  — комментарий
 * H  — источник (какой «якорь»/страница откуда отправили)
 *
 * САМЫЙ ПРОСТОЙ ВАРИАНТ — ИСПОЛЬЗОВАТЬ ИМЕНА СТОЛБЦОВ КАК В ШАПКЕ ТАБЛИЦЫ:
 * если у вас уже есть таблица с шапкой (Дата, Имя, Телефон, ...), укажите
 * здесь эти же имена и код автоматически запишет данные в нужные столбцы.
 */

const CONFIG = {
  /* Название листа. Если такого листа нет — он будет создан. */
  SHEET_NAME: 'Заявки',

  /* Названия колонок (первая строка). Меняйте под свою таблицу. */
  HEADERS: ['Дата заявки', 'Имя', 'Телефон', 'Эффект', 'Желаемая дата', 'Время', 'Комментарий', 'Источник'],
};

/**
 * Принимает POST-запрос от формы:
 * тело запроса — JSON:
 * {
 *   "name": "Анна",
 *   "phone": "+7 900 123-45-67",
 *   "effect": "2D",
 *   "date": "2026-09-25",
 *   "time": "14:00",
 *   "comment": "хочу естественный эффект",
 *   "source": "catalog-2d"   // необязательное поле
 * }
 */
function doPost(e) {
  try {
    const body = typeof e.postData.contents === 'string'
      ? JSON.parse(e.postData.contents)
      : {};
    appendLead(body);
    return jsonAnswer({ ok: true });
  } catch (err) {
    return jsonAnswer({ ok: false, error: String(err) });
  }
}

/** GET — проверка, что веб-приложение живо (откройте URL в браузере). */
function doGet() {
  const html = '<h2>Сбор заявок работает</h2>'
    + '<p>Отправьте POST-запрос с JSON-данными заявки. '
    + 'Подробности — в комментариях к коду скрипта.</p>';
  return HtmlService.createHtmlOutput(html);
}

function appendLead(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    /* Проект не привязан к таблице (Web App развёрнут отдельно).
       Привяжите таблицу: откройте её через Сервис -> Скриптовый редактор,
       либо вставьте SpreadsheetApp.openById("ID_таблицы") ниже. */
    throw new Error("Проект не привязан к Google-таблице: откройте таблицу через Сервис -> Скриптовый редактор и переразверните Web App.");
  }

  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    /* Лист с таким именем ещё не существует — создаём с шапкой. */
    try {
      sheet = ss.insertSheet(CONFIG.SHEET_NAME);
      sheet.getRange(1, 1, 1, CONFIG.HEADERS.length)
        .setValues([CONFIG.HEADERS])
        .setFontWeight('bold');
    } catch (e) {
      /* Если и создать не вышло — пишем в первый существующий лист. */
      sheet = ss.getSheets()[0] || null;
    }
  }
  if (!sheet) {
    throw new Error("Не найден лист для записи: " + CONFIG.SHEET_NAME);
  }

  /* Убеждаемся, что в таблице есть шапка (на случай, если лист создали вручную). */
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const hasHeader = headerRow.some(String);

  const values = {
    'Дата заявки':   new Date(),
    'Имя':           clean(data.name),
    'Телефон':       clean(data.phone),
    'Эффект':        clean(data.effect),
    'Желаемая дата': clean(data.date),
    'Время':         clean(data.time),
    'Комментарий':   clean(data.comment),
    'Источник':      clean(data.source),
  };

  let rowValues;
  if (hasHeader) {
    /* Привязка по именам колонок из шапки */
    rowValues = CONFIG.HEADERS.map((h) => {
      const hit = headerRow.indexOf(h);
      if (hit !== -1) return values[h] || '';
      return values[h];
    });
    /* Если колонка не совпала с шапкой — просто кладём в конец ряда по порядку */
    if (rowValues.length < CONFIG.HEADERS.length) {
      rowValues = CONFIG.HEADERS.map((h) => values[h] || '');
    }
  } else {
    rowValues = CONFIG.HEADERS.map((h) => values[h] || '');
  }

  sheet.appendRow(rowValues);

  /* Телефон всегда пишем как ТЕКСТ ('@'), чтобы Google Sheets не превращал
     номер в число и не рушил формат (например "89001234567" -> 8.9E+10,
     потеря "+7"/нулей). */
  const phoneCol = headerRow.indexOf('Телефон') + 1;
  if (phoneCol > 0){
    const lastRow = sheet.getLastRow();
    const phoneCell = sheet.getRange(lastRow, phoneCol);
    phoneCell.setNumberFormat('@').setValue(values['Телефон']);
  }
}

function clean(v) {
  if (v === undefined || v === null) return '';
  return String(v).trim();
}

function jsonAnswer(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ============================================================
   ПРИМЕР ПОДКЛЮЧЕНИЯ К ЛЕНДИНГУ (frontend, вставить в index.html)
   В обработчике submit формы замените frontend-логику на:

   const APP_SCRIPT_URL = "https://script.google.com/macros/s/XXXXXXXX/exec";

   fetch(APP_SCRIPT_URL, {
     method: "POST",
     headers: { "Content-Type": "text/plain;charset=utf-8" },
     body: JSON.stringify({
       name: fields.name.value.trim(),
       phone: fields.phone.value.trim(),
       effect: fields.effect.value,
       date: fields.date.value,
       time: fields.time.value,
       comment: fields.comment.value.trim(),
       source: form.dataset.source || ""
     })
   })
   .then(r => r.json())
   .then(res => {
     if (res.ok) {
       formView.hidden = true;
       successView.hidden = false;
     } else {
       alert("Не удалось отправить заявку. Попробуйте ещё раз.");
     }
   })
   .catch(() => alert("Не удалось отправить заявку. Попробуйте ещё раз."));

   ВАЖНО: отправляйте Content-Type: text/plain, чтобы не срабатывал
   предварительный CORS-запрос OPTIONS. Apps Script сам разберёт JSON.
   ============================================================ */