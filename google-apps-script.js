// ════════════════════════════════════════════
//  語音代辦 Agent — Google Apps Script
//  階段一：語音輸入 → Google 行事曆 / Tasks
//  階段二：Notion 更新 → 同步到 Google
// ════════════════════════════════════════════

// ── 設定區（填入你的值）──────────────────────
const NOTION_TOKEN    = 'ntn_你的Token';
const NOTION_DB_ID    = '你的DatabaseID';
const SHOPPING_LIST   = '🛒 購物清單';
const MISC_LIST       = '📌 雜項';
// ────────────────────────────────────────────


// ════════════════════════════════════════════
//  階段一：接收 App 語音輸入
// ════════════════════════════════════════════

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.type === 'calendar') return respond(createCalendarEvent(data));
    if (data.type === 'shopping' || data.type === 'misc') return respond(createTask(data));
    return respond({ success: false, error: 'Unknown type: ' + data.type });
  } catch (err) {
    return respond({ success: false, error: err.message });
  }
}

// ── Google 行事曆 ────────────────────────────
function createCalendarEvent(data) {
  const cal     = CalendarApp.getDefaultCalendar();
  const dateStr = data.date || new Date().toISOString().split('T')[0];
  let event;

  if (data.time) {
    const start = new Date(dateStr + 'T' + data.time + ':00');
    const end   = new Date(start.getTime() + 60 * 60 * 1000);
    event = cal.createEvent(data.name, start, end);
  } else {
    event = cal.createAllDayEvent(data.name, new Date(dateStr + 'T00:00:00'));
  }

  // 同步寫入 Notion
  writeToNotion({ name: data.name, type: 'calendar', date: dateStr, time: data.time || null, priority: data.priority || 'none' });

  return { success: true, type: 'calendar', event_id: event.getId(), title: event.getTitle() };
}

// ── Google Tasks ─────────────────────────────
function createTask(data) {
  const listName = data.type === 'shopping' ? SHOPPING_LIST : MISC_LIST;
  const listId   = getOrCreateTaskList(listName).id;
  const task     = Tasks.newTask();
  task.title     = data.name;
  if (data.date) task.due = new Date(data.date + 'T00:00:00').toISOString();
  const created  = Tasks.Tasks.insert(task, listId);

  // 同步寫入 Notion
  writeToNotion({ name: data.name, type: data.type, date: data.date || null, priority: data.priority || 'none' });

  return { success: true, type: data.type, task_id: created.id, title: created.title };
}

function getOrCreateTaskList(name) {
  const lists = (Tasks.Tasklists.list().items || []);
  const found  = lists.find(l => l.title === name);
  if (found) return found;
  const newList  = Tasks.newTaskList();
  newList.title  = name;
  return Tasks.Tasklists.insert(newList);
}


// ════════════════════════════════════════════
//  Notion：寫入資料庫
// ════════════════════════════════════════════

function writeToNotion(data) {
  const priorityMap = { high: 'High', medium: 'Medium', low: 'Low', none: null };
  const properties  = {
    "Name":   { title: [{ text: { content: data.name } }] },
    "Status": { status: { name: 'Not started' } }
  };

  if (data.date) {
    properties["Due Date"] = {
      date: { start: data.date + (data.time ? 'T' + data.time + ':00' : '') }
    };
  }
  if (priorityMap[data.priority]) {
    properties["Priority"] = { select: { name: priorityMap[data.priority] } };
  }

  const res = UrlFetchApp.fetch('https://api.notion.com/v1/pages', {
    method: 'post',
    headers: {
      'Authorization': 'Bearer ' + NOTION_TOKEN,
      'Notion-Version': '2022-06-28',
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify({ parent: { database_id: NOTION_DB_ID.replace(/-/g, '') }, properties }),
    muteHttpExceptions: true
  });

  const result = JSON.parse(res.getContentText());
  return { success: !!result.id, notion_id: result.id };
}


// ════════════════════════════════════════════
//  階段二：Notion 更新 → 同步到 Google
//  在 Apps Script 執行 setupTrigger() 一次即可
// ════════════════════════════════════════════

function syncFromNotion() {
  const since = new Date(Date.now() - 10 * 60 * 1000).toISOString();

  const res = UrlFetchApp.fetch(
    'https://api.notion.com/v1/databases/' + NOTION_DB_ID.replace(/-/g,'') + '/query', {
    method: 'post',
    headers: {
      'Authorization': 'Bearer ' + NOTION_TOKEN,
      'Notion-Version': '2022-06-28',
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify({
      filter: { timestamp: 'last_edited_time', last_edited_time: { on_or_after: since } },
      sorts:  [{ timestamp: 'last_edited_time', direction: 'descending' }],
      page_size: 20
    }),
    muteHttpExceptions: true
  });

  const data = JSON.parse(res.getContentText());
  if (!data.results) { Logger.log('Notion error: ' + res.getContentText()); return; }

  data.results.forEach(page => processNotionPage(page));
  Logger.log('同步完成，處理 ' + data.results.length + ' 筆');
}

function processNotionPage(page) {
  const props    = page.properties;
  const name     = props['Name']?.title?.[0]?.plain_text || '';
  const status   = props['Status']?.status?.name || '';
  const dateVal  = props['Due Date']?.date?.start || null;
  const priority = props['Priority']?.select?.name?.toLowerCase() || 'none';

  if (!name) return;
  if (status === 'Done' || status === 'Complete') return; // 已完成不重複同步

  // 避免重複：用 PropertiesService 記錄已處理的 page ID
  const store    = PropertiesService.getScriptProperties();
  const syncedKey = 'synced_' + page.id;
  if (store.getProperty(syncedKey)) return;
  store.setProperty(syncedKey, new Date().toISOString());

  const hasTime  = dateVal && dateVal.includes('T');
  const dateOnly = dateVal ? dateVal.split('T')[0] : null;
  const timeOnly = hasTime ? dateVal.split('T')[1].substring(0, 5) : null;

  if (hasTime && dateOnly) {
    // 有時間 → 行事曆
    const cal   = CalendarApp.getDefaultCalendar();
    const start = new Date(dateOnly + 'T' + timeOnly + ':00');
    const end   = new Date(start.getTime() + 60 * 60 * 1000);
    cal.createEvent(name, start, end);
    Logger.log('行事曆 ✅ ' + name);
  } else {
    // 無時間 → Tasks 雜項
    const listId = getOrCreateTaskList(MISC_LIST).id;
    const task   = Tasks.newTask();
    task.title   = name;
    if (dateOnly) task.due = new Date(dateOnly + 'T00:00:00').toISOString();
    Tasks.Tasks.insert(task, listId);
    Logger.log('Tasks ✅ ' + name);
  }
}

// 執行一次設定定時觸發器
function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'syncFromNotion') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncFromNotion').timeBased().everyMinutes(5).create();
  Logger.log('✅ 已設定：每 5 分鐘同步一次 Notion → Google');
}


// ════════════════════════════════════════════
//  測試用（在編輯器直接執行）
// ════════════════════════════════════════════

function test() {
  Logger.log(JSON.stringify(createCalendarEvent({
    name: '測試會議', date: new Date().toISOString().split('T')[0], time: '14:00', priority: 'high'
  })));
  Logger.log(JSON.stringify(createTask({ name: '買牛奶', type: 'shopping' })));
  Logger.log(JSON.stringify(createTask({ name: '回信給客戶', type: 'misc', priority: 'medium' })));
}

function respond(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
