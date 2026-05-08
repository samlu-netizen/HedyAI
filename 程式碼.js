/**
 * ==========================================
 * Kanban Board - Multi-User Task Dispatch System
 * ==========================================
 */

// Webhook 驗證用 Token 已從靜態變數改為使用 PropertiesService 動態取得
// 預設 Token 為 "hedy-ai-secret-token-123"

// 使用者提供的雲端硬碟共用資料夾 ID
const TEAM_FOLDER_ID = "1muc66GWMoq2HIhepU_OPUty8NOAdw_Jy";
const HEDY_API_BASE_URL = "https://api.hedy.bot";
const HEDY_HEADER_SYNC_LIMIT = 5;
const ADMIN_MEMBER_NAME = 'Admin';

// 預設團隊成員；實際名單會在首次執行後存入 Script Properties，並可由前端註冊新增。
const DEFAULT_TEAM_MEMBERS = [ADMIN_MEMBER_NAME, "Alice", "Bob", "Charlie", "Dave", "Eve"];
const TEAM_MEMBERS_PROPERTY_KEY = 'REGISTERED_TEAM_MEMBERS';
const DISABLED_TEAM_MEMBERS_PROPERTY_KEY = 'DISABLED_TEAM_MEMBERS';

const TASK_HEADERS = ['id', 'title', 'description', 'status', 'priority', 'dueDate', 'createdAt', 'updatedAt', 'source', 'assigner', 'assignee', 'meetingId', 'meetingTitle', 'externalSessionId', 'topic', 'topicLabel', 'memo'];
const MEETING_HEADERS = ['id', 'source', 'eventType', 'meetingTitle', 'meetingDate', 'externalSessionId', 'topic', 'status', 'transcript', 'summaryMarkdown', 'recapMarkdown', 'actionItemsMarkdown', 'decisionsMarkdown', 'nextStepsMarkdown', 'rawPayload', 'actionItemCount', 'createdAt', 'updatedAt', 'createdBy'];
const HIGHLIGHT_HEADERS = ['id', 'highlightId', 'sessionId', 'meetingId', 'timestamp', 'title', 'rawQuote', 'cleanedQuote', 'mainIdea', 'aiInsight', 'rawPayload', 'createdAt', 'updatedAt'];
const LOG_HEADERS = ['timestamp', 'status', 'message', 'payload'];
const HEDY_WEBHOOK_EVENTS = ['session.ended', 'session.exported', 'highlight.created', 'todo.exported'];
const MAX_LOG_ROWS = 20;
const MAX_MEETING_ROWS = 100;

let teamMapCache_ = null;
let spreadsheetCache_ = {};

/**
 * 處理 HTTP GET 請求，回傳 Web App 頁面
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Kanban 任務派發系統')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 取得團隊的試算表對照表
 * 首次執行時會掃描並自動在資料夾內幫所有人建立專屬的試算表
 */
function getTeamMap() {
  if (teamMapCache_) {
    return teamMapCache_;
  }

  const props = PropertiesService.getScriptProperties();
  let mapString = props.getProperty('TEAM_MAP');
  let map = mapString ? JSON.parse(mapString) : {};
  
  let needsUpdate = false;
  const members = getRegisteredTeamMembers_();
  
  if (!hasCompleteTeamMap_(map)) {
    const folder = DriveApp.getFolderById(TEAM_FOLDER_ID);
    for (const member of members) {
      if (map[member]) continue;

      // 找尋同名檔案 Kanban_{Member}
      const files = folder.searchFiles(`title = 'Kanban_${member}'`);
      if (files.hasNext()) {
        map[member] = files.next().getId();
      } else {
        // 沒有的話就自動建立
        const newSheet = SpreadsheetApp.create(`Kanban_${member}`);
        const fileId = newSheet.getId();
        
        // 將檔案移動到指定資料夾 (V8 寫法)
        const file = DriveApp.getFileById(fileId);
        file.moveTo(folder);
        
        map[member] = fileId;
        
        // 初始化表格與欄位
        const taskSheet = newSheet.getSheets()[0];
        taskSheet.setName('Tasks');
        ensureSheetHeaders(taskSheet, TASK_HEADERS);

        const meetingSheet = newSheet.insertSheet('Meetings');
        ensureSheetHeaders(meetingSheet, MEETING_HEADERS);

        const highlightSheet = newSheet.insertSheet('Highlights');
        ensureSheetHeaders(highlightSheet, HIGHLIGHT_HEADERS);
        
        const logSheet = newSheet.insertSheet('Logs');
        ensureSheetHeaders(logSheet, LOG_HEADERS);
      }
      needsUpdate = true;
    }
  }
  
  if (needsUpdate) {
    props.setProperty('TEAM_MAP', JSON.stringify(map));
  }

  teamMapCache_ = map;
  return map;
}

function hasCompleteTeamMap_(map) {
  return getRegisteredTeamMembers_().every(member => Boolean(map && map[member]));
}

function ensureSheetHeaders(sheet, headers) {
  const range = sheet.getRange(1, 1, 1, headers.length);
  range.setValues([headers]);
  range.setFontWeight('bold').setBackground('#f3f4f6');
  sheet.setFrozenRows(1);
}

function ensureMemberSchema(spreadsheet) {
  const tasks = spreadsheet.getSheetByName('Tasks') || spreadsheet.insertSheet('Tasks');
  ensureSheetHeadersIfNeeded_(tasks, TASK_HEADERS);

  const meetings = spreadsheet.getSheetByName('Meetings') || spreadsheet.insertSheet('Meetings');
  ensureSheetHeadersIfNeeded_(meetings, MEETING_HEADERS);

  const highlights = spreadsheet.getSheetByName('Highlights') || spreadsheet.insertSheet('Highlights');
  ensureSheetHeadersIfNeeded_(highlights, HIGHLIGHT_HEADERS);

  const logs = spreadsheet.getSheetByName('Logs') || spreadsheet.insertSheet('Logs');
  ensureSheetHeadersIfNeeded_(logs, LOG_HEADERS);
}

function ensureSheetHeadersIfNeeded_(sheet, headers) {
  const currentWidth = Math.max(sheet.getLastColumn(), headers.length);
  const existing = currentWidth > 0
    ? sheet.getRange(1, 1, 1, currentWidth).getDisplayValues()[0]
    : [];
  const matches = headers.every((header, index) => existing[index] === header);

  if (!matches) {
    ensureSheetHeaders(sheet, headers);
  }
}

/**
 * 供給前端，取得團隊人員名單
 */
function getTeamMembers() {
  return getActiveTeamMembers_();
}

function getRegisteredTeamMembers_() {
  const props = PropertiesService.getScriptProperties();
  const stored = props.getProperty(TEAM_MEMBERS_PROPERTY_KEY);
  let members = [];

  if (stored) {
    try {
      members = JSON.parse(stored);
    } catch (err) {
      members = [];
    }
  }

  if (!Array.isArray(members) || members.length === 0) {
    members = DEFAULT_TEAM_MEMBERS.slice();
  }
  if (!members.some(member => normalizeMemberName_(member).toLowerCase() === ADMIN_MEMBER_NAME.toLowerCase())) {
    members.unshift(ADMIN_MEMBER_NAME);
  }

  const normalizedMembers = [];
  members.forEach((member) => {
    const normalized = normalizeMemberName_(member);
    if (normalized && !normalizedMembers.some(existing => existing.toLowerCase() === normalized.toLowerCase())) {
      normalizedMembers.push(normalized);
    }
  });

  if (JSON.stringify(normalizedMembers) !== stored) {
    props.setProperty(TEAM_MEMBERS_PROPERTY_KEY, JSON.stringify(normalizedMembers));
  }

  return normalizedMembers;
}

function getDisabledTeamMembers_() {
  const props = PropertiesService.getScriptProperties();
  const stored = props.getProperty(DISABLED_TEAM_MEMBERS_PROPERTY_KEY);
  let disabledMembers = [];

  if (stored) {
    try {
      disabledMembers = JSON.parse(stored);
    } catch (err) {
      disabledMembers = [];
    }
  }

  const registeredMembers = getRegisteredTeamMembers_();
  const normalizedDisabledMembers = [];
  (Array.isArray(disabledMembers) ? disabledMembers : []).forEach((member) => {
    const normalized = normalizeMemberName_(member);
    if (
      normalized &&
      registeredMembers.includes(normalized) &&
      normalized !== ADMIN_MEMBER_NAME &&
      normalizedDisabledMembers.indexOf(normalized) < 0
    ) {
      normalizedDisabledMembers.push(normalized);
    }
  });

  if (JSON.stringify(normalizedDisabledMembers) !== stored) {
    props.setProperty(DISABLED_TEAM_MEMBERS_PROPERTY_KEY, JSON.stringify(normalizedDisabledMembers));
  }

  return normalizedDisabledMembers;
}

function getActiveTeamMembers_() {
  const disabledMembers = getDisabledTeamMembers_();
  return getRegisteredTeamMembers_().filter(member => disabledMembers.indexOf(member) < 0);
}

function setDisabledTeamMembers_(members) {
  const registeredMembers = getRegisteredTeamMembers_();
  const normalizedMembers = [];
  (Array.isArray(members) ? members : []).forEach((member) => {
    const normalized = normalizeMemberName_(member);
    if (
      normalized &&
      registeredMembers.includes(normalized) &&
      normalizedMembers.indexOf(normalized) < 0
    ) {
      normalizedMembers.push(normalized);
    }
  });
  PropertiesService.getScriptProperties().setProperty(DISABLED_TEAM_MEMBERS_PROPERTY_KEY, JSON.stringify(normalizedMembers));
  return normalizedMembers;
}

function isAdminUser_(user) {
  return normalizeMemberName_(user) === ADMIN_MEMBER_NAME;
}

function assertAdminUser_(user) {
  if (!isAdminUser_(user)) {
    throw new Error('只有 Admin 可以管理團隊成員');
  }
}

function getTeamMemberAdminData(currentUser) {
  const disabledMembers = getDisabledTeamMembers_();
  const members = getRegisteredTeamMembers_();
  return {
    members: members.map(member => ({
      name: member,
      disabled: disabledMembers.indexOf(member) >= 0
    })),
    activeMembers: members.filter(member => disabledMembers.indexOf(member) < 0),
    disabledMembers: disabledMembers,
    adminUser: ADMIN_MEMBER_NAME,
    canManageTeamMembers: isAdminUser_(currentUser)
  };
}

function normalizeMemberName_(name) {
  return String(name || '')
    .replace(/\s+/g, ' ')
    .trim();
}

function validateNewMemberName_(name) {
  const normalized = normalizeMemberName_(name);
  if (!normalized) {
    throw new Error('請輸入使用者名稱');
  }

  if (normalized.length > 40) {
    throw new Error('使用者名稱不能超過 40 個字');
  }

  if (/[\\/?*\[\]:]/.test(normalized)) {
    throw new Error('使用者名稱不能包含 \\ / ? * [ ] :');
  }

  return normalized;
}

function registerTeamMember(name) {
  const normalizedName = validateNewMemberName_(name);
  const props = PropertiesService.getScriptProperties();
  const members = getRegisteredTeamMembers_();
  const existing = members.find(member => member.toLowerCase() === normalizedName.toLowerCase());

  if (existing) {
    const disabledMembers = getDisabledTeamMembers_();
    const isDisabled = disabledMembers.indexOf(existing) >= 0;
    if (isDisabled) {
      setDisabledTeamMembers_(disabledMembers.filter(member => member !== existing));
    }
    return {
      created: false,
      restored: isDisabled,
      member: existing,
      members: getActiveTeamMembers_()
    };
  }

  members.push(normalizedName);
  props.setProperty(TEAM_MEMBERS_PROPERTY_KEY, JSON.stringify(members));
  teamMapCache_ = null;

  try {
    getUserSpreadsheet_(normalizedName, true);
  } catch (err) {
    const rolledBackMembers = members.filter(member => member !== normalizedName);
    props.setProperty(TEAM_MEMBERS_PROPERTY_KEY, JSON.stringify(rolledBackMembers));
    teamMapCache_ = null;
    throw err;
  }

  return {
    created: true,
    member: normalizedName,
    members: getActiveTeamMembers_()
  };
}

function disableTeamMember(name, currentUser) {
  const normalizedName = normalizeMemberName_(name);
  const normalizedCurrentUser = normalizeMemberName_(currentUser);
  assertAdminUser_(normalizedCurrentUser);
  const members = getRegisteredTeamMembers_();
  if (!members.includes(normalizedName)) {
    throw new Error('找不到要停用的使用者');
  }
  if (normalizedName === ADMIN_MEMBER_NAME) {
    throw new Error('不能停用 Admin 使用者');
  }
  if (normalizedName === normalizedCurrentUser) {
    throw new Error('不能停用目前登入的使用者');
  }

  const disabledMembers = getDisabledTeamMembers_();
  if (disabledMembers.indexOf(normalizedName) < 0) {
    disabledMembers.push(normalizedName);
    setDisabledTeamMembers_(disabledMembers);
  }

  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(getHedySyncCursorKey_(normalizedName));
  return getTeamMemberAdminData(normalizedCurrentUser);
}

function restoreTeamMember(name, currentUser) {
  const normalizedCurrentUser = normalizeMemberName_(currentUser);
  assertAdminUser_(normalizedCurrentUser);
  const normalizedName = normalizeMemberName_(name);
  const members = getRegisteredTeamMembers_();
  if (!members.includes(normalizedName)) {
    throw new Error('找不到要恢復的使用者');
  }

  setDisabledTeamMembers_(getDisabledTeamMembers_().filter(member => member !== normalizedName));
  getUserSpreadsheet_(normalizedName, true);
  return getTeamMemberAdminData(normalizedCurrentUser);
}

function getInitialAppData(preferredUser) {
  const members = getActiveTeamMembers_();
  const requestedUser = String(preferredUser || '').trim();
  const currentUser = members.includes(requestedUser) ? requestedUser : (members[0] || '');
  const tasks = currentUser ? getTasks(currentUser) : [];

  return {
    members: members,
    currentUser: currentUser,
    tasks: tasks
  };
}

function resolveWebhookUser_(params) {
  const user = String((params && params.user) || '').trim();
  if (!user) {
    return '';
  }
  const activeMembers = getActiveTeamMembers_();
  return activeMembers.includes(user) ? user : '';
}

function setCurrentUser(user) {
  const normalizedUser = normalizeMemberName_(user);
  if (!getActiveTeamMembers_().includes(normalizedUser)) {
    throw new Error('請先選擇有效的使用者');
  }
  PropertiesService.getUserProperties().setProperty('LAST_USER', normalizedUser);
  return true;
}

function getUserSpreadsheet_(user, ensureSchema) {
  const normalizedUser = String(user || '').trim();
  if (!normalizedUser) {
    throw new Error('未指定使用者');
  }

  const map = getTeamMap();
  if (!map[normalizedUser]) {
    throw new Error(`找不到使用者 ${normalizedUser} 的表單`);
  }

  if (spreadsheetCache_[normalizedUser]) {
    if (ensureSchema) {
      ensureMemberSchema(spreadsheetCache_[normalizedUser]);
    }
    return spreadsheetCache_[normalizedUser];
  }

  const ss = SpreadsheetApp.openById(map[normalizedUser]);
  if (ensureSchema) {
    ensureMemberSchema(ss);
  }
  spreadsheetCache_[normalizedUser] = ss;
  return ss;
}

/**
 * 取得與當前使用者相關的任務 (無論是負責的或派出去的)
 */
function getTasks(currentUser) {
  if (!currentUser) return [];
  const ss = getUserSpreadsheet_(currentUser, false);
  const sheet = ss.getSheetByName('Tasks');
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow <= 1 || lastColumn === 0) return [];

  const data = sheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues();

  const headers = data.shift();
  return data.map(row => {
    let task = {};
    headers.forEach((header, i) => task[header] = row[i]);
    return normalizeLegacyTaskForUi(task);
  });
}

function getAssignedTasks(currentUser) {
  const assigner = String(currentUser || '').trim();
  if (!assigner) return [];

  const assignedTasks = [];
  getActiveTeamMembers_().forEach((member) => {
    if (member === assigner) return;

    try {
      const ss = getUserSpreadsheet_(member, false);
      const sheet = ss.getSheetByName('Tasks');
      if (!sheet) return;

      const lastRow = sheet.getLastRow();
      const lastColumn = sheet.getLastColumn();
      if (lastRow <= 1 || lastColumn === 0) return;

      const data = sheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues();
      const headers = data.shift();
      const assignerIndex = headers.indexOf('assigner');
      const assigneeIndex = headers.indexOf('assignee');
      if (assignerIndex < 0 || assigneeIndex < 0) return;

      data.forEach((row) => {
        if (String(row[assignerIndex] || '').trim() !== assigner) return;
        if (String(row[assigneeIndex] || '').trim() === assigner) return;

        const task = {};
        headers.forEach((header, i) => task[header] = row[i]);
        assignedTasks.push(normalizeLegacyTaskForUi(task));
      });
    } catch (err) {
      // Ignore one member sheet failure so the tracker can still show other assignees.
    }
  });

  return assignedTasks;
}

/**
 * 新增/派發任務
 */
function addTask(task) {
  // 任務寫入的目標永遠是負責人(Assignee)
  const targetMember = task.assignee || task.assigner; 
  if (targetMember === ADMIN_MEMBER_NAME) {
    throw new Error('Admin 僅供系統管理使用，不能被指派任務');
  }
  if (!getActiveTeamMembers_().includes(targetMember)) {
    throw new Error('無法指派給已停用或不存在的使用者');
  }
  const ss = getUserSpreadsheet_(targetMember, true);
  const sheet = ss.getSheetByName('Tasks');

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const record = buildTaskRecord_(task);
  const row = headers.map((header) => record[header] || '');

  sheet.appendRow(row);
  return getTaskById(record.id, targetMember);
}

function buildTaskRecord_(task) {
  const now = new Date().toISOString();
  const assigner = task.assigner || '';
  const assignee = task.assignee || '';
  const isAssignedToAnotherMember = assigner && assignee && assigner !== assignee;
  return {
    id: task.id || Utilities.getUuid(),
    title: task.title,
    description: task.description || '',
    status: isAssignedToAnotherMember ? 'todo' : (task.status || 'todo'),
    priority: task.priority || 'medium',
    dueDate: task.dueDate || '',
    createdAt: task.createdAt || now,
    updatedAt: task.updatedAt || now,
    source: task.source || 'manual',
    assigner: assigner,
    assignee: assignee,
    assignerName: assigner,
    assigneeName: assignee,
    assignerEmail: '',
    assigneeEmail: '',
    meetingId: task.meetingId || '',
    meetingTitle: task.meetingTitle || '',
    externalSessionId: task.externalSessionId || '',
    topic: task.topic || '',
    topicLabel: task.topicLabel || '',
    memo: task.memo || ''
  };
}

function appendTasksBatch_(targetUser, tasks) {
  const safeTasks = Array.isArray(tasks) ? tasks.filter(Boolean) : [];
  if (safeTasks.length === 0) return 0;

  const ss = getUserSpreadsheet_(targetUser, true);
  const sheet = ss.getSheetByName('Tasks');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const records = safeTasks.map(buildTaskRecord_);
  const rows = records.map(record => headers.map((header) => record[header] || ''));

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  return rows.length;
}

function updateTaskRowFields_(sheet, rowIndex, headers, fieldMap) {
  Object.keys(fieldMap).forEach((field) => {
    const colIndex = headers.indexOf(field);
    if (colIndex >= 0) {
      sheet.getRange(rowIndex, colIndex + 1).setValue(fieldMap[field]);
    }
  });
}

/**
 * 輔助函數：在全表尋找任務實體
 */
function findTaskCtx(id, user) {
  const ss = getUserSpreadsheet_(user, false);
  const sheet = ss.getSheetByName('Tasks');
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      return { ss: ss, sheet: sheet, rowIndex: i + 1, data: data[i], headers: data[0] };
    }
  }
  return null;
}

function findTaskCtxAcrossTeam_(id) {
  const taskId = String(id || '').trim();
  if (!taskId) return null;

  const members = getRegisteredTeamMembers_();
  for (let i = 0; i < members.length; i++) {
    const member = members[i];
    try {
      const ctx = findTaskCtx(taskId, member);
      if (ctx) {
        return ctx;
      }
    } catch (err) {
      // Ignore a single broken member sheet and continue searching.
    }
  }

  return null;
}

function findTaskCtxWithHints_(id, users) {
  const taskId = String(id || '').trim();
  if (!taskId) return null;

  const members = getRegisteredTeamMembers_();
  const checkedUsers = [];
  (Array.isArray(users) ? users : []).forEach((user) => {
    const normalizedUser = String(user || '').trim();
    if (normalizedUser && members.includes(normalizedUser) && checkedUsers.indexOf(normalizedUser) < 0) {
      checkedUsers.push(normalizedUser);
    }
  });

  for (let i = 0; i < checkedUsers.length; i++) {
    try {
      const ctx = findTaskCtx(taskId, checkedUsers[i]);
      if (ctx) return ctx;
    } catch (err) {
      // Fall through to the broader lookup below.
    }
  }

  return findTaskCtxAcrossTeam_(taskId);
}

function findTaskCtxsByExternalSessionId_(externalSessionId, user) {
  const normalizedSessionId = String(externalSessionId || '').trim();
  if (!normalizedSessionId) return [];

  const ss = getUserSpreadsheet_(user, false);
  const sheet = ss.getSheetByName('Tasks');
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow <= 1 || lastColumn === 0) return [];

  const data = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  const headers = data[0];
  const sessionIndex = headers.indexOf('externalSessionId');
  if (sessionIndex < 0) return [];

  const matches = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][sessionIndex] || '').trim() === normalizedSessionId) {
      matches.push({
        ss: ss,
        sheet: sheet,
        rowIndex: i + 1,
        data: data[i],
        headers: headers
      });
    }
  }

  return matches;
}

/**
 * 更新任務內容
 */
function updateTask(taskData) {
  let ctx = findTaskCtxWithHints_(taskData.id, [
    taskData.previousAssignee,
    taskData.currentUser,
    taskData.assignee
  ]);
  if (!ctx) throw new Error("Task not found");

  if (ctx.headers.indexOf('memo') < 0) {
    ensureMemberSchema(ctx.ss);
    ctx = findTaskCtxWithHints_(taskData.id, [
      taskData.previousAssignee,
      taskData.currentUser,
      taskData.assignee
    ]);
    if (!ctx) throw new Error("Task not found");
  }
  
  const now = new Date().toISOString();
  const h = ctx.headers;
  const oldAssignee = ctx.data[h.indexOf('assignee')];
  const existingMeetingId = h.indexOf('meetingId') >= 0 ? (ctx.data[h.indexOf('meetingId')] || '') : '';
  const existingMeetingTitle = h.indexOf('meetingTitle') >= 0 ? (ctx.data[h.indexOf('meetingTitle')] || '') : '';
  const existingExternalSessionId = h.indexOf('externalSessionId') >= 0 ? (ctx.data[h.indexOf('externalSessionId')] || '') : '';
  const existingTopic = h.indexOf('topic') >= 0 ? (ctx.data[h.indexOf('topic')] || '') : '';
  const existingTopicLabel = h.indexOf('topicLabel') >= 0 ? (ctx.data[h.indexOf('topicLabel')] || '') : '';
  const existingMemo = h.indexOf('memo') >= 0 ? (ctx.data[h.indexOf('memo')] || '') : '';
  
  // 如果負責人改變了，要把舊表的任務刪除，寫入新表 (重新派發)
  if (taskData.assignee && taskData.assignee !== oldAssignee) {
      copyMeetingToUserIfNeeded_(oldAssignee, taskData.assignee, existingMeetingId);
      ctx.sheet.deleteRow(ctx.rowIndex);
      return addTask({
          ...taskData,
          status: 'todo',
          assigner: taskData.assigner || ctx.data[h.indexOf('assigner')],
          createdAt: ctx.data[h.indexOf('createdAt')],
          source: ctx.data[h.indexOf('source')],
          meetingId: existingMeetingId,
          meetingTitle: existingMeetingTitle,
          externalSessionId: existingExternalSessionId,
          topic: existingTopic,
          topicLabel: existingTopicLabel,
          memo: taskData.memo || existingMemo
      });
  }
  
  const nextData = ctx.data.slice();
  const updates = {
    title: taskData.title,
    description: taskData.description || '',
    status: taskData.status,
    priority: taskData.priority,
    dueDate: taskData.dueDate || '',
    memo: taskData.memo || '',
    updatedAt: now
  };
  Object.keys(updates).forEach((field) => {
    const colIndex = h.indexOf(field);
    if (colIndex >= 0) {
      nextData[colIndex] = updates[field];
    }
  });
  ctx.sheet.getRange(ctx.rowIndex, 1, 1, h.length).setValues([nextData]);
  ctx.data = nextData;

  const updatedTask = {};
  h.forEach((header, index) => {
    updatedTask[header] = nextData[index];
  });
  return normalizeLegacyTaskForUi(updatedTask);
}

function updateTaskMemo(id, memo, currentUser, taskOwnerUser) {
  let ctx = findTaskCtxWithHints_(id, [
    taskOwnerUser,
    currentUser
  ]);
  if (!ctx) throw new Error("Task not found");

  if (ctx.headers.indexOf('memo') < 0) {
    ensureMemberSchema(ctx.ss);
    ctx = findTaskCtxWithHints_(id, [
      taskOwnerUser,
      currentUser
    ]);
    if (!ctx) throw new Error("Task not found");
  }

  const effectiveUser = String(currentUser || '').trim();
  const h = ctx.headers;
  const assigner = h.indexOf('assigner') >= 0 ? String(ctx.data[h.indexOf('assigner')] || '').trim() : '';
  const assignee = h.indexOf('assignee') >= 0 ? String(ctx.data[h.indexOf('assignee')] || '').trim() : '';
  if (!effectiveUser || (assigner !== effectiveUser && assignee !== effectiveUser)) {
    throw new Error('只有派發者或負責人可以更新任務記事');
  }

  const now = new Date().toISOString();
  const nextMemo = String(memo || '').trim();
  updateTaskRowFields_(ctx.sheet, ctx.rowIndex, h, {
    memo: nextMemo,
    updatedAt: now
  });

  const memoIndex = h.indexOf('memo');
  const updatedAtIndex = h.indexOf('updatedAt');
  if (memoIndex >= 0) ctx.data[memoIndex] = nextMemo;
  if (updatedAtIndex >= 0) ctx.data[updatedAtIndex] = now;

  const updatedTask = {};
  h.forEach((header, index) => {
    updatedTask[header] = ctx.data[index];
  });
  return normalizeLegacyTaskForUi(updatedTask);
}

/**
 * 更新狀態 (拖曳)
 */
function updateTaskStatus(id, newStatus, currentUser) {
  const effectiveUser = String(currentUser || '').trim() || PropertiesService.getUserProperties().getProperty('LAST_USER') || '';
  if (!effectiveUser) {
    throw new Error('未指定目前使用者，無法更新任務狀態');
  }

  const task = getTaskById(id, effectiveUser);
  const ctx = task ? findTaskCtxAcrossTeam_(id) : null;
  if (!ctx) return false;
  
  const now = new Date().toISOString();
  ctx.sheet.getRange(ctx.rowIndex, ctx.headers.indexOf('status') + 1).setValue(newStatus);
  ctx.sheet.getRange(ctx.rowIndex, ctx.headers.indexOf('updatedAt') + 1).setValue(now);
  return true;
}

/**
 * 刪除任務
 */
function deleteTask(id) {
  const task = getTaskById(id, PropertiesService.getUserProperties().getProperty('LAST_USER') || '');
  const ctx = task ? findTaskCtxAcrossTeam_(id) : null;
  if (!ctx) throw new Error("Task not found");
  ctx.sheet.deleteRow(ctx.rowIndex);
  return true;
}

function getTaskById(id, user) {
  let ctx = null;
  const lookupUser = String(user || '').trim();
  if (lookupUser) {
    try {
      ctx = findTaskCtx(id, lookupUser);
    } catch (err) {
      ctx = null;
    }
  }
  if (!ctx) {
    ctx = findTaskCtxAcrossTeam_(id);
  }
  if (!ctx) return null;
  
  let task = {};
  ctx.headers.forEach((header, index) => {
    task[header] = ctx.data[index];
  });
  return normalizeLegacyTaskForUi(task);
}

function normalizeLegacyTaskForUi(task) {
  const normalized = Object.assign({}, task || {});
  const assigner = normalized.assigner || normalized.assignerName || normalized.assignerEmail || '';
  const assignee = normalized.assignee || normalized.assigneeName || normalized.assigneeEmail || '';

  // Compatibility for rows that were written by older code into newer schemas:
  // meetingId/topic may accidentally contain assigner/assignee display values.
  if (!normalized.assigner && !normalized.assignerName && !normalized.assignerEmail && normalized.meetingId && !looksLikeUuid_(normalized.meetingId)) {
    normalized.assigner = normalized.meetingId;
  } else {
    normalized.assigner = assigner;
  }

  if (!normalized.assignee && !normalized.assigneeName && !normalized.assigneeEmail && normalized.topic && normalized.topic !== '未分類') {
    normalized.assignee = normalized.topic;
  } else {
    normalized.assignee = assignee;
  }

  return normalized;
}

function looksLikeUuid_(value) {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(String(value || '').trim());
}

function getUserEmailByName_(memberName) {
  const normalized = String(memberName || '').trim().toLowerCase();
  if (!normalized) return '';

  try {
    const usersSheet = SpreadsheetApp.openById(getMasterSpreadsheetId_()).getSheetByName('Users');
    if (!usersSheet || usersSheet.getLastRow() <= 1) return '';
    const rows = usersSheet.getDataRange().getDisplayValues();
    const headers = rows.shift();
    const displayNameIndex = headers.indexOf('displayName');
    const legacyNameIndex = headers.indexOf('legacyMemberName');
    const emailIndex = headers.indexOf('email');
    if (emailIndex < 0) return '';

    const match = rows.find((row) => {
      const displayName = displayNameIndex >= 0 ? String(row[displayNameIndex] || '').trim().toLowerCase() : '';
      const legacyName = legacyNameIndex >= 0 ? String(row[legacyNameIndex] || '').trim().toLowerCase() : '';
      return displayName === normalized || legacyName === normalized;
    });
    return match ? String(match[emailIndex] || '').trim() : '';
  } catch (err) {
    return '';
  }
}

function getMasterSpreadsheetId_() {
  const props = PropertiesService.getScriptProperties();
  const storedId = props.getProperty('MASTER_SPREADSHEET_ID');
  if (storedId) return storedId;
  const folder = DriveApp.getFolderById(TEAM_FOLDER_ID);
  const files = folder.searchFiles("title = 'Kanban_Master'");
  if (files.hasNext()) {
    const id = files.next().getId();
    props.setProperty('MASTER_SPREADSHEET_ID', id);
    return id;
  }
  return '';
}

/**
 * 取得紀錄 (個人)
 */
function getLogs(user) {
  if (!user) return [];
  const ss = getUserSpreadsheet_(user, false);
  const logSheet = ss.getSheetByName('Logs');
  if(!logSheet) return [];
  
  const lastRow = logSheet.getLastRow();
  const lastColumn = logSheet.getLastColumn();
  if (lastRow <= 1 || lastColumn === 0) return [];

  const startRow = Math.max(2, lastRow - MAX_LOG_ROWS + 1);
  const numRows = lastRow - startRow + 1;
  const data = logSheet.getRange(startRow, 1, numRows, lastColumn).getDisplayValues();
  return data.reverse().map(r => ({
    timestamp: r[0],
    status: r[1],
    message: r[2],
    payload: String(r[3] || ''),
    payloadPreview: buildPayloadPreview_(r[3]),
    payloadIsJson: isJsonString_(r[3])
  }));
}

function logEvent(user, status, message, payload) {
  if(!user) return;
  const ss = getUserSpreadsheet_(user, true);
  const logSheet = ss.getSheetByName('Logs');
  if(logSheet) {
      logSheet.appendRow([
        new Date().toISOString(), 
        status, 
        message, 
        stringifyPayload_(payload)
      ]);
  }
}

function stringifyPayload_(payload) {
  if (payload === null || payload === undefined) {
    return '';
  }

  if (typeof payload === 'string') {
    return payload;
  }

  try {
    return JSON.stringify(payload);
  } catch (err) {
    return String(payload);
  }
}

function isJsonString_(value) {
  const text = String(value || '').trim();
  if (!text) return false;

  try {
    JSON.parse(text);
    return true;
  } catch (err) {
    return false;
  }
}

function buildPayloadPreview_(value) {
  const text = String(value || '').trim();
  if (!text) return '無 payload';

  if (isJsonString_(text)) {
    try {
      const parsed = JSON.parse(text);
      if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
        const keys = Object.keys(parsed).slice(0, 5);
        return keys.length > 0 ? keys.join(', ') : 'JSON payload';
      }
      if (Array.isArray(parsed)) {
        return `JSON array (${parsed.length})`;
      }
    } catch (err) {
      // Fallback to plain-text preview below.
    }
  }

  return text.length > 120 ? `${text.slice(0, 117)}...` : text;
}

function getMeetings(currentUser) {
  if (!currentUser) return [];
  const ss = getUserSpreadsheet_(currentUser, false);
  const sheet = ss.getSheetByName('Meetings');
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow <= 1 || lastColumn === 0) return [];

  const startRow = Math.max(2, lastRow - MAX_MEETING_ROWS + 1);
  const numRows = lastRow - startRow + 1;
  const data = sheet.getRange(startRow, 1, numRows, lastColumn).getDisplayValues();

  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  return data
    .map(row => {
      const meeting = {};
      headers.forEach((header, i) => {
        meeting[header] = row[i];
      });
      return normalizeMeetingForUi_(meeting);
    })
    .sort((a, b) => String(b.meetingDate || '').localeCompare(String(a.meetingDate || '')));
}

function getMeetingById(meetingId, currentUser) {
  const ctx = findMeetingCtx_(meetingId, currentUser);
  if (!ctx) return null;

  const meeting = {};
  ctx.headers.forEach((header, index) => {
    meeting[header] = ctx.data[index];
  });
  return normalizeMeetingForUi_(meeting);
}

function getMeetingRecordById_(meetingId, user) {
  const ctx = findMeetingCtx_(meetingId, user);
  if (!ctx) return null;

  const meeting = {};
  ctx.headers.forEach((header, index) => {
    meeting[header] = ctx.data[index];
  });
  return meeting;
}

function findMeetingByExternalSessionId_(externalSessionId, user) {
  const ctx = findMeetingCtxByExternalSessionId_(externalSessionId, user);
  if (!ctx) return null;

  const meeting = {};
  ctx.headers.forEach((header, index) => {
    meeting[header] = ctx.data[index];
  });
  return normalizeMeetingForUi_(meeting);
}

function findMeetingCtxByExternalSessionId_(externalSessionId, user) {
  const normalizedSessionId = String(externalSessionId || '').trim();
  if (!normalizedSessionId) return null;

  const ss = getUserSpreadsheet_(user, false);
  const sheet = ss.getSheetByName('Meetings');
  if (!sheet) return null;

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow <= 1 || lastColumn === 0) return null;

  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  const sessionIndex = headers.indexOf('externalSessionId');
  if (sessionIndex < 0) return null;

  const data = sheet.getRange(2, 1, lastRow - 1, lastColumn).getDisplayValues();
  for (let i = data.length - 1; i >= 0; i--) {
    if (String(data[i][sessionIndex] || '').trim() === normalizedSessionId) {
      return {
        ss: ss,
        sheet: sheet,
        rowIndex: i + 2,
        data: data[i],
        headers: headers
      };
    }
  }

  return null;
}

function findMeetingCtx_(meetingId, user) {
  const ss = getUserSpreadsheet_(user, false);
  const sheet = ss.getSheetByName('Meetings');
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === meetingId) {
      return { ss: ss, sheet: sheet, rowIndex: i + 1, data: data[i], headers: data[0] };
    }
  }
  return null;
}

function normalizeMeetingForUi_(meeting) {
  const normalized = Object.assign({}, meeting || {});
  normalized.topicLabel = normalized.topic || '未分類';
  normalized.actionItemCount = Number(normalized.actionItemCount || 0);
  return normalized;
}

function normalizeActionItem_(item) {
  if (!item) return null;
  if (typeof item === 'string') {
    return {
      title: item.trim(),
      details: ''
    };
  }

  return {
    title: String(item.task || item.title || '').trim(),
    details: String(item.details || item.content || '').trim()
  };
}

function parseActionItemsFromSummary_(summaryMarkdown) {
  const text = String(summaryMarkdown || '');
  const match = text.match(/## 行動項目([\s\S]*?)(\n## |\s*$)/);
  if (!match) return [];

  return match[1]
    .split('\n')
    .map(line => line.trim())
    .filter(line => line.startsWith('- '))
    .map(line => {
      const cleaned = line
        .replace(/^-+\s*/, '')
        .replace(/\*\*/g, '')
        .trim();
      return {
        title: cleaned,
        details: ''
      };
    })
    .filter(item => item.title);
}

function parseBulletLines_(text) {
  return String(text || '')
    .split('\n')
    .map(line => line.trim())
    .filter(line => line.startsWith('- ') || line.startsWith('• '))
    .map(line => line.replace(/^[-•]\s*/, '').trim())
    .filter(Boolean);
}

function extractMarkdownSection_(text, heading) {
  const escapedHeading = String(heading || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const pattern = new RegExp(`##\\s*${escapedHeading}\\s*([\\s\\S]*?)(\\n##\\s|$)`);
  const match = String(text || '').match(pattern);
  return match ? match[1].trim() : '';
}

function extractActionItemsFromPayload_(payload) {
  const rawItems = []
    .concat(Array.isArray(payload.actionItems) ? payload.actionItems : [])
    .concat(Array.isArray(payload.todos) ? payload.todos : [])
    .concat(Array.isArray(payload.user_todos) ? payload.user_todos : [])
    .concat(Array.isArray(payload.userTodos) ? payload.userTodos : [])
    .concat(Array.isArray(payload['待辦事項']) ? payload['待辦事項'] : [])
    .concat(payload.todo ? [payload.todo] : []);

  const explicitItems = rawItems
    .map(normalizeActionItem_)
    .filter(item => item && item.title);

  if (explicitItems.length > 0) {
    return explicitItems;
  }

  const meetingMinutes = String(payload.meeting_minutes || '');
  const actionSection = extractMarkdownSection_(meetingMinutes, '行動項目');
  const fallbackSection = actionSection || extractMarkdownSection_(meetingMinutes, '待辦事項');
  const bulletItems = parseBulletLines_(fallbackSection).map(line => ({ title: line, details: '' }));
  if (bulletItems.length > 0) {
    return bulletItems;
  }

  const summaryMarkdown = String(payload.recap || payload.summary || '');
  return parseActionItemsFromSummary_(summaryMarkdown);
}

function buildMeetingDetailFields_(payload) {
  const meetingMinutes = String(payload.meeting_minutes || '');
  const recap = String(payload.recap || payload.summary || '');
  const actionItems = extractActionItemsFromPayload_(payload);
  const decisionsMarkdown = extractMarkdownSection_(meetingMinutes, '決議事項');
  const nextStepsMarkdown = extractMarkdownSection_(meetingMinutes, '下一步');

  const summaryBlocks = [recap, meetingMinutes].filter(Boolean);
  const summaryMarkdown = summaryBlocks.join('\n\n').trim();
  const actionItemsMarkdown = actionItems.map(item => `- ${item.title}`).join('\n');

  return {
    summaryMarkdown: summaryMarkdown,
    recapMarkdown: recap,
    actionItemsMarkdown: actionItemsMarkdown,
    decisionsMarkdown: decisionsMarkdown,
    nextStepsMarkdown: nextStepsMarkdown
  };
}

function getHedyTopicLabel_(payload) {
  if (!payload) return '';
  if (payload.topicName) return String(payload.topicName || '').trim();
  if (payload.topic && typeof payload.topic === 'object') {
    return String(payload.topic.name || payload.topic.id || '').trim();
  }
  return String(payload.topic || payload.topicId || '').trim();
}

function getHedyMeetingStatus_(payload, eventType) {
  if (payload && payload.status) return String(payload.status || '').trim();
  if (eventType === 'session.created') return 'recording';
  if (eventType === 'session.ended') return 'completed';
  if (eventType === 'session.exported') return 'exported';
  return getHedyTopicLabel_(payload) ? 'classified' : 'unclassified';
}

function getHedyMeetingDate_(payload, eventType) {
  if (!payload) return '';
  if (eventType === 'session.exported' && payload.exportedAt) return payload.exportedAt;
  return payload.meetingDate || payload.date || payload.startTime || payload.endTime || payload.exportedAt || payload.createdAt || '';
}

function normalizeHedyWebhookEvent_(payload) {
  const envelope = payload && typeof payload === 'object' ? payload : {};
  const hasOfficialData = envelope.data && typeof envelope.data === 'object' && !Array.isArray(envelope.data);
  const data = hasOfficialData ? envelope.data : envelope;
  const eventType = String(envelope.event || envelope.type || envelope.event_type || data.event || data.type || data.event_type || 'unknown').trim() || 'unknown';

  return {
    eventType: eventType,
    data: data || {},
    rawPayload: envelope
  };
}

function createMeetingRecord_(payload, eventType, createdBy, rawPayload) {
  const now = new Date().toISOString();
  const detailFields = buildMeetingDetailFields_(payload);
  const fallbackItems = extractActionItemsFromPayload_(payload);
  const topicLabel = getHedyTopicLabel_(payload);
  const meetingDate = getHedyMeetingDate_(payload, eventType) || now;
  const sessionId = String(payload.sessionId || payload.id || '').trim();
  const title = payload.title || payload.meetingTitle || `Hedy 會議匯入 (${eventType})`;

  return {
    id: Utilities.getUuid(),
    source: 'hedyai_webhook',
    eventType: eventType,
    meetingTitle: title,
    meetingDate: meetingDate,
    externalSessionId: sessionId,
    topic: topicLabel,
    status: getHedyMeetingStatus_(payload, eventType),
    transcript: payload.cleaned_transcript || payload.transcript || '',
    summaryMarkdown: detailFields.summaryMarkdown,
    recapMarkdown: detailFields.recapMarkdown,
    actionItemsMarkdown: detailFields.actionItemsMarkdown,
    decisionsMarkdown: detailFields.decisionsMarkdown,
    nextStepsMarkdown: detailFields.nextStepsMarkdown,
    rawPayload: JSON.stringify(rawPayload || payload || {}),
    actionItemCount: fallbackItems.length,
    createdAt: now,
    updatedAt: now,
    createdBy: createdBy || 'Hedy AI'
  };
}

function buildTasksFromWebhookPayload_(payload, targetUser, meetingRecord) {
  const parsedItems = extractActionItemsFromPayload_(payload);
  const summaryMarkdown = meetingRecord.summaryMarkdown || payload.meeting_minutes || payload.recap || payload.summary || '';

  if (parsedItems.length === 0) {
    return [{
      id: Utilities.getUuid(),
      title: `回顧會議：${meetingRecord.meetingTitle}`,
      description: payload.transcript || summaryMarkdown || '會議已匯入，但沒有解析到明確代辦事項。',
      status: 'todo',
      priority: 'medium',
      dueDate: '',
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      source: 'hedyai_webhook',
      assigner: 'Hedy AI',
      assignee: targetUser,
      meetingId: meetingRecord.id,
      topic: meetingRecord.topic || '',
      topicLabel: meetingRecord.topic || ''
    }];
  }

  return parsedItems.map(item => ({
    id: Utilities.getUuid(),
    title: item.title,
    description: item.details || `來自會議：${meetingRecord.meetingTitle}`,
    status: payload.completed ? 'done' : 'todo',
    priority: 'medium',
    dueDate: '',
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    source: 'hedyai_webhook',
    assigner: 'Hedy AI',
    assignee: targetUser,
    meetingId: meetingRecord.id,
    topic: meetingRecord.topic || '',
    topicLabel: meetingRecord.topic || ''
  }));
}

function appendMeetingRecord_(user, meetingRecord) {
  if (!user) throw new Error('無效的會議寫入對象');
  const ss = getUserSpreadsheet_(user, true);
  ss.getSheetByName('Meetings').appendRow(MEETING_HEADERS.map(header => meetingRecord[header] || ''));
}

function upsertMeetingRecord_(user, meetingRecord) {
  if (!user) throw new Error('無效的會議寫入對象');
  getUserSpreadsheet_(user, true);
  const existingCtx = meetingRecord.externalSessionId
    ? findMeetingCtxByExternalSessionId_(meetingRecord.externalSessionId, user)
    : null;
  const ctx = existingCtx || findMeetingCtx_(meetingRecord.id, user);

  if (!ctx) {
    appendMeetingRecord_(user, meetingRecord);
    return {
      created: true,
      meeting: meetingRecord
    };
  }

  const merged = Object.assign({}, meetingRecord, {
    id: ctx.data[ctx.headers.indexOf('id')] || meetingRecord.id,
    createdAt: ctx.data[ctx.headers.indexOf('createdAt')] || meetingRecord.createdAt,
    updatedAt: new Date().toISOString()
  });

  const row = ctx.headers.map(header => merged[header] || '');
  ctx.sheet.getRange(ctx.rowIndex, 1, 1, row.length).setValues([row]);
  return {
    created: false,
    meeting: merged
  };
}

function findHighlightCtx_(highlightId, user) {
  const normalizedHighlightId = String(highlightId || '').trim();
  if (!normalizedHighlightId) return null;

  const ss = getUserSpreadsheet_(user, false);
  const sheet = ss.getSheetByName('Highlights');
  if (!sheet) return null;

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow <= 1 || lastColumn === 0) return null;

  const data = sheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues();
  const headers = data[0];
  const highlightIndex = headers.indexOf('highlightId');
  if (highlightIndex < 0) return null;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][highlightIndex] || '').trim() === normalizedHighlightId) {
      return {
        ss: ss,
        sheet: sheet,
        rowIndex: i + 1,
        data: data[i],
        headers: headers
      };
    }
  }

  return null;
}

function createHighlightRecord_(payload, targetUser, rawPayload) {
  const now = new Date().toISOString();
  const highlightId = String(payload.highlightId || payload.id || '').trim() || Utilities.getUuid();
  const sessionId = String(payload.sessionId || '').trim();
  const meeting = sessionId ? findMeetingByExternalSessionId_(sessionId, targetUser) : null;

  return {
    id: Utilities.getUuid(),
    highlightId: highlightId,
    sessionId: sessionId,
    meetingId: meeting ? String(meeting.id || '').trim() : '',
    timestamp: payload.timestamp || now,
    title: payload.title || 'Hedy Highlight',
    rawQuote: payload.rawQuote || '',
    cleanedQuote: payload.cleanedQuote || '',
    mainIdea: payload.mainIdea || '',
    aiInsight: payload.aiInsight || '',
    rawPayload: JSON.stringify(rawPayload || payload || {}),
    createdAt: now,
    updatedAt: now
  };
}

function upsertHighlightRecord_(user, highlightRecord) {
  if (!user) throw new Error('無效的亮點寫入對象');
  const ss = getUserSpreadsheet_(user, true);
  const sheet = ss.getSheetByName('Highlights');
  const ctx = findHighlightCtx_(highlightRecord.highlightId, user);

  if (!ctx) {
    sheet.appendRow(HIGHLIGHT_HEADERS.map(header => highlightRecord[header] || ''));
    return {
      created: true,
      highlight: highlightRecord
    };
  }

  const merged = Object.assign({}, highlightRecord, {
    id: ctx.data[ctx.headers.indexOf('id')] || highlightRecord.id,
    createdAt: ctx.data[ctx.headers.indexOf('createdAt')] || highlightRecord.createdAt,
    updatedAt: new Date().toISOString()
  });
  const row = ctx.headers.map(header => merged[header] || '');
  ctx.sheet.getRange(ctx.rowIndex, 1, 1, row.length).setValues([row]);
  return {
    created: false,
    highlight: merged
  };
}

function upsertEmbeddedHighlights_(user, payload, meetingRecord, rawPayload) {
  const highlights = Array.isArray(payload && payload.highlights) ? payload.highlights : [];
  if (highlights.length === 0) return 0;

  let count = 0;
  highlights.forEach((highlight) => {
    if (!highlight) return;
    const enrichedHighlight = Object.assign({}, highlight, {
      sessionId: highlight.sessionId || meetingRecord.externalSessionId || ''
    });
    const record = createHighlightRecord_(enrichedHighlight, user, rawPayload || payload);
    if (!record.meetingId && meetingRecord.id) {
      record.meetingId = meetingRecord.id;
    }
    upsertHighlightRecord_(user, record);
    count++;
  });

  return count;
}

function copyMeetingToUserIfNeeded_(sourceUser, targetUser, meetingId) {
  const normalizedMeetingId = String(meetingId || '').trim();
  if (!sourceUser || !targetUser || !normalizedMeetingId || sourceUser === targetUser) {
    return null;
  }

  if (findMeetingCtx_(normalizedMeetingId, targetUser)) {
    return getMeetingRecordById_(normalizedMeetingId, targetUser);
  }

  const sourceMeeting = getMeetingRecordById_(normalizedMeetingId, sourceUser);
  if (!sourceMeeting) {
    return null;
  }

  appendMeetingRecord_(targetUser, sourceMeeting);
  return sourceMeeting;
}

function linkExistingTasksToMeeting_(user, meetingRecord) {
  const sessionId = String((meetingRecord && meetingRecord.externalSessionId) || '').trim();
  if (!user || !sessionId) return 0;

  const matches = findTaskCtxsByExternalSessionId_(sessionId, user);
  let linkedCount = 0;
  matches.forEach((ctx) => {
    updateTaskRowFields_(ctx.sheet, ctx.rowIndex, ctx.headers, {
      meetingId: meetingRecord.id || '',
      meetingTitle: meetingRecord.meetingTitle || '',
      topic: meetingRecord.topic || '',
      topicLabel: meetingRecord.topic || ''
    });
    linkedCount++;
  });

  return linkedCount;
}

function createTaskFromTodoExport_(payload, targetUser) {
  const meeting = findMeetingByExternalSessionId_(payload.sessionId, targetUser);
  const topicLabel = getHedyTopicLabel_(payload);

  return {
    id: payload.id || Utilities.getUuid(),
    title: String(payload.text || '').trim() || 'Hedy AI 匯出的待辦事項',
    description: payload.sessionId ? `來自 Hedy Session: ${payload.sessionId}` : '來自 Hedy AI 待辦事項匯出',
    status: 'todo',
    priority: 'medium',
    dueDate: String(payload.dueDate || '').trim(),
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    source: 'hedyai_todo_export',
    assigner: 'Hedy AI',
    assignee: targetUser,
    meetingId: meeting ? String(meeting.id || '').trim() : '',
    meetingTitle: meeting ? String(meeting.meetingTitle || '').trim() : '',
    externalSessionId: String(payload.sessionId || '').trim(),
    topic: meeting ? String(meeting.topic || '').trim() : topicLabel,
    topicLabel: meeting ? String(meeting.topicLabel || '').trim() : topicLabel
  };
}

function upsertTodoExportTask_(payload, targetUser) {
  const todoId = String(payload.id || '').trim();
  if (todoId) {
    const existing = findTaskCtx(todoId, targetUser);
    if (existing) {
      return {
        created: false,
        taskId: todoId
      };
    }
  }

  const task = createTaskFromTodoExport_(payload, targetUser);
  addTask(task);
  return {
    created: true,
    taskId: task.id
  };
}

function getHedyApiKey_(overrideKey) {
  const providedKey = String(overrideKey || '').trim();
  if (providedKey) return providedKey;
  return String(PropertiesService.getScriptProperties().getProperty('HEDY_API_KEY') || '').trim();
}

function hedyApiRequest_(path, options) {
  const requestOptions = options || {};
  const apiKey = getHedyApiKey_(requestOptions.apiKey);
  if (!apiKey) {
    throw new Error('尚未設定 Hedy API Key');
  }

  const fetchOptions = {
    method: requestOptions.method || 'get',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + apiKey
    },
    muteHttpExceptions: true
  };
  if (requestOptions.payload) {
    fetchOptions.payload = JSON.stringify(requestOptions.payload);
  }

  const response = UrlFetchApp.fetch(HEDY_API_BASE_URL + path, fetchOptions);

  const statusCode = response.getResponseCode();
  const bodyText = response.getContentText();
  const body = bodyText ? JSON.parse(bodyText) : {};

  if (statusCode < 200 || statusCode >= 300) {
    const message = body && body.error && body.error.message ? body.error.message : bodyText;
    throw new Error(`Hedy API ${statusCode}: ${message || 'request failed'}`);
  }

  return body;
}

function unwrapHedyList_(response) {
  if (Array.isArray(response)) return response;
  if (response && Array.isArray(response.data)) return response.data;
  return [];
}

function readIndexedSheet_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow <= 0 || lastColumn === 0) {
    return {
      sheet: sheet,
      headers: [],
      rows: []
    };
  }

  const values = sheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues();
  return {
    sheet: sheet,
    headers: values[0] || [],
    rows: values.slice(1).map((row, index) => ({
      sheet: sheet,
      headers: values[0] || [],
      data: row,
      rowIndex: index + 2
    }))
  };
}

function recordFromCtx_(ctx) {
  const record = {};
  if (!ctx) return record;
  ctx.headers.forEach((header, index) => {
    record[header] = ctx.data[index];
  });
  return record;
}

function createHedySyncIndex_(targetUser) {
  const ss = getUserSpreadsheet_(targetUser, true);
  const meetings = readIndexedSheet_(ss.getSheetByName('Meetings'));
  const tasks = readIndexedSheet_(ss.getSheetByName('Tasks'));
  const highlights = readIndexedSheet_(ss.getSheetByName('Highlights'));

  const meetingsById = {};
  const meetingsByExternalSessionId = {};
  const meetingIdIndex = meetings.headers.indexOf('id');
  const meetingSessionIndex = meetings.headers.indexOf('externalSessionId');
  meetings.rows.forEach((ctx) => {
    const id = meetingIdIndex >= 0 ? String(ctx.data[meetingIdIndex] || '').trim() : '';
    const sessionId = meetingSessionIndex >= 0 ? String(ctx.data[meetingSessionIndex] || '').trim() : '';
    if (id) meetingsById[id] = ctx;
    if (sessionId) meetingsByExternalSessionId[sessionId] = ctx;
  });

  const tasksById = {};
  const taskIdIndex = tasks.headers.indexOf('id');
  tasks.rows.forEach((ctx) => {
    const id = taskIdIndex >= 0 ? String(ctx.data[taskIdIndex] || '').trim() : '';
    if (id) tasksById[id] = ctx;
  });

  const highlightsByHighlightId = {};
  const highlightIdIndex = highlights.headers.indexOf('highlightId');
  highlights.rows.forEach((ctx) => {
    const id = highlightIdIndex >= 0 ? String(ctx.data[highlightIdIndex] || '').trim() : '';
    if (id) highlightsByHighlightId[id] = ctx;
  });

  return {
    ss: ss,
    meetings: Object.assign(meetings, {
      byId: meetingsById,
      byExternalSessionId: meetingsByExternalSessionId
    }),
    tasks: Object.assign(tasks, {
      byId: tasksById
    }),
    highlights: Object.assign(highlights, {
      byHighlightId: highlightsByHighlightId
    })
  };
}

function getHedySyncCursorKey_(targetUser) {
  return `HEDY_SESSIONS_CURSOR_${targetUser}`;
}

function upsertMeetingRecordWithIndex_(syncIndex, meetingRecord) {
  const meetings = syncIndex.meetings;
  const existingCtx = meetingRecord.externalSessionId
    ? meetings.byExternalSessionId[String(meetingRecord.externalSessionId || '').trim()]
    : null;
  const ctx = existingCtx || meetings.byId[String(meetingRecord.id || '').trim()];

  if (!ctx) {
    const row = meetings.headers.map(header => meetingRecord[header] || '');
    meetings.sheet.appendRow(row);
    const nextCtx = {
      sheet: meetings.sheet,
      headers: meetings.headers,
      data: row,
      rowIndex: meetings.sheet.getLastRow()
    };
    if (meetingRecord.id) meetings.byId[String(meetingRecord.id).trim()] = nextCtx;
    if (meetingRecord.externalSessionId) meetings.byExternalSessionId[String(meetingRecord.externalSessionId).trim()] = nextCtx;
    return {
      created: true,
      meeting: meetingRecord
    };
  }

  const merged = Object.assign({}, meetingRecord, {
    id: ctx.data[ctx.headers.indexOf('id')] || meetingRecord.id,
    createdAt: ctx.data[ctx.headers.indexOf('createdAt')] || meetingRecord.createdAt,
    updatedAt: new Date().toISOString()
  });
  const row = ctx.headers.map(header => merged[header] || '');
  ctx.sheet.getRange(ctx.rowIndex, 1, 1, row.length).setValues([row]);
  ctx.data = row;
  if (merged.id) meetings.byId[String(merged.id).trim()] = ctx;
  if (merged.externalSessionId) meetings.byExternalSessionId[String(merged.externalSessionId).trim()] = ctx;
  return {
    created: false,
    meeting: merged
  };
}

function createHighlightRecordFromMeeting_(payload, meetingRecord, rawPayload) {
  const now = new Date().toISOString();
  const highlightId = String(payload.highlightId || payload.id || '').trim() || Utilities.getUuid();
  const sessionId = String(payload.sessionId || meetingRecord.externalSessionId || '').trim();

  return {
    id: Utilities.getUuid(),
    highlightId: highlightId,
    sessionId: sessionId,
    meetingId: meetingRecord.id || '',
    timestamp: payload.timestamp || now,
    title: payload.title || 'Hedy Highlight',
    rawQuote: payload.rawQuote || '',
    cleanedQuote: payload.cleanedQuote || '',
    mainIdea: payload.mainIdea || '',
    aiInsight: payload.aiInsight || '',
    rawPayload: JSON.stringify(rawPayload || payload || {}),
    createdAt: now,
    updatedAt: now
  };
}

function upsertHighlightRecordWithIndex_(syncIndex, highlightRecord) {
  const highlights = syncIndex.highlights;
  const highlightId = String(highlightRecord.highlightId || '').trim();
  const ctx = highlightId ? highlights.byHighlightId[highlightId] : null;

  if (!ctx) {
    const row = highlights.headers.map(header => highlightRecord[header] || '');
    highlights.sheet.appendRow(row);
    const nextCtx = {
      sheet: highlights.sheet,
      headers: highlights.headers,
      data: row,
      rowIndex: highlights.sheet.getLastRow()
    };
    if (highlightId) highlights.byHighlightId[highlightId] = nextCtx;
    return {
      created: true,
      highlight: highlightRecord
    };
  }

  const merged = Object.assign({}, highlightRecord, {
    id: ctx.data[ctx.headers.indexOf('id')] || highlightRecord.id,
    createdAt: ctx.data[ctx.headers.indexOf('createdAt')] || highlightRecord.createdAt,
    updatedAt: new Date().toISOString()
  });
  const row = ctx.headers.map(header => merged[header] || '');
  ctx.sheet.getRange(ctx.rowIndex, 1, 1, row.length).setValues([row]);
  ctx.data = row;
  if (merged.highlightId) highlights.byHighlightId[String(merged.highlightId).trim()] = ctx;
  return {
    created: false,
    highlight: merged
  };
}

function upsertEmbeddedHighlightsWithIndex_(syncIndex, payload, meetingRecord, rawPayload) {
  const highlights = Array.isArray(payload && payload.highlights) ? payload.highlights : [];
  if (highlights.length === 0) return 0;

  let count = 0;
  highlights.forEach((highlight) => {
    if (!highlight) return;
    const enrichedHighlight = Object.assign({}, highlight, {
      sessionId: highlight.sessionId || meetingRecord.externalSessionId || ''
    });
    const record = createHighlightRecordFromMeeting_(enrichedHighlight, meetingRecord, rawPayload || payload);
    upsertHighlightRecordWithIndex_(syncIndex, record);
    count++;
  });

  return count;
}

function createTaskFromTodoExportWithIndex_(payload, targetUser, syncIndex) {
  const sessionId = String(payload.sessionId || '').trim();
  const meetingCtx = sessionId ? syncIndex.meetings.byExternalSessionId[sessionId] : null;
  const meeting = meetingCtx ? normalizeMeetingForUi_(recordFromCtx_(meetingCtx)) : null;
  const topicLabel = getHedyTopicLabel_(payload);

  return {
    id: payload.id || Utilities.getUuid(),
    title: String(payload.text || '').trim() || 'Hedy AI 匯出的待辦事項',
    description: sessionId ? `來自 Hedy Session: ${sessionId}` : '來自 Hedy AI 待辦事項匯出',
    status: 'todo',
    priority: 'medium',
    dueDate: String(payload.dueDate || '').trim(),
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    source: 'hedyai_todo_export',
    assigner: 'Hedy AI',
    assignee: targetUser,
    meetingId: meeting ? String(meeting.id || '').trim() : '',
    meetingTitle: meeting ? String(meeting.meetingTitle || '').trim() : '',
    externalSessionId: sessionId,
    topic: meeting ? String(meeting.topic || '').trim() : topicLabel,
    topicLabel: meeting ? String(meeting.topicLabel || '').trim() : topicLabel
  };
}

function upsertTodoExportTaskWithIndex_(payload, targetUser, syncIndex) {
  const todoId = String(payload.id || '').trim();
  if (todoId && syncIndex.tasks.byId[todoId]) {
    return {
      created: false,
      taskId: todoId
    };
  }

  const task = createTaskFromTodoExportWithIndex_(payload, targetUser, syncIndex);
  const record = buildTaskRecord_(task);
  const row = syncIndex.tasks.headers.map((header) => record[header] || '');
  syncIndex.tasks.sheet.appendRow(row);
  const ctx = {
    sheet: syncIndex.tasks.sheet,
    headers: syncIndex.tasks.headers,
    data: row,
    rowIndex: syncIndex.tasks.sheet.getLastRow()
  };
  if (record.id) syncIndex.tasks.byId[String(record.id).trim()] = ctx;
  return {
    created: true,
    taskId: record.id
  };
}

function syncRecentHedySessions_(targetUser, limit, afterCursor, syncIndex) {
  const safeLimit = Math.max(1, Math.min(Number(limit || 20), 50));
  const normalizedCursor = String(afterCursor || '').trim();
  const path = `/sessions?limit=${safeLimit}` + (normalizedCursor ? `&after=${encodeURIComponent(normalizedCursor)}` : '');
  const listResponse = hedyApiRequest_(path);
  const sessions = unwrapHedyList_(listResponse);
  let created = 0;
  let updated = 0;
  let highlights = 0;
  const sessionIds = [];

  sessions.forEach((session) => {
    const sessionId = String((session && session.sessionId) || '').trim();
    if (!sessionId) return;
    sessionIds.push(sessionId);

    const detail = hedyApiRequest_(`/sessions/${encodeURIComponent(sessionId)}`);
    const meetingRecord = createMeetingRecord_(detail, 'api.session.sync', 'Hedy API', detail);
    meetingRecord.source = 'hedyai_api';
    const result = syncIndex
      ? upsertMeetingRecordWithIndex_(syncIndex, meetingRecord)
      : upsertMeetingRecord_(targetUser, meetingRecord);
    if (result.created) created++;
    else updated++;
    highlights += syncIndex
      ? upsertEmbeddedHighlightsWithIndex_(syncIndex, detail, result.meeting, detail)
      : upsertEmbeddedHighlights_(targetUser, detail, result.meeting, detail);
  });

  return {
    created: created,
    updated: updated,
    highlights: highlights,
    sessionIds: sessionIds,
    pagination: listResponse && listResponse.pagination ? listResponse.pagination : null
  };
}

function syncHedyTodos_(targetUser) {
  const response = hedyApiRequest_('/todos');
  const todos = unwrapHedyList_(response);
  let created = 0;
  let skipped = 0;

  todos.forEach((todo) => {
    if (!todo) return;
    const result = upsertTodoExportTask_(todo, targetUser);
    if (result.created) created++;
    else skipped++;
  });

  return {
    created: created,
    skipped: skipped
  };
}

function syncHedyTodosForSessions_(targetUser, sessionIds, syncIndex) {
  const uniqueSessionIds = Array.from(new Set((sessionIds || [])
    .map((sessionId) => String(sessionId || '').trim())
    .filter(Boolean)));
  let created = 0;
  let skipped = 0;
  let checked = 0;

  uniqueSessionIds.forEach((sessionId) => {
    const response = hedyApiRequest_(`/sessions/${encodeURIComponent(sessionId)}/todos`);
    const todos = unwrapHedyList_(response);
    checked += todos.length;

    todos.forEach((todo) => {
      if (!todo) return;
      const todoWithSession = Object.assign({}, todo, {
        sessionId: todo.sessionId || sessionId
      });
      const result = syncIndex
        ? upsertTodoExportTaskWithIndex_(todoWithSession, targetUser, syncIndex)
        : upsertTodoExportTask_(todoWithSession, targetUser);
      if (result.created) created++;
      else skipped++;
    });
  });

  return {
    created: created,
    skipped: skipped,
    checked: checked,
    sessions: uniqueSessionIds.length
  };
}


function syncHedyHighlights_(targetUser, limit) {
  const safeLimit = Math.max(1, Math.min(Number(limit || 50), 100));
  const listResponse = hedyApiRequest_(`/highlights?limit=${safeLimit}`);
  const highlights = unwrapHedyList_(listResponse);
  let created = 0;
  let updated = 0;

  highlights.forEach((highlight) => {
    if (!highlight) return;
    const highlightId = String(highlight.highlightId || highlight.id || '').trim();
    const detail = highlightId
      ? hedyApiRequest_(`/highlights/${encodeURIComponent(highlightId)}`)
      : highlight;
    const record = createHighlightRecord_(detail, targetUser, detail);
    const result = upsertHighlightRecord_(targetUser, record);
    if (result.created) created++;
    else updated++;
  });

  return {
    created: created,
    updated: updated
  };
}

function syncHedyApiData(currentUser) {
  const targetUser = String(currentUser || '').trim();
  if (!getActiveTeamMembers_().includes(targetUser)) {
    throw new Error('未指定有效使用者，無法同步 Hedy API');
  }

  const sessions = syncRecentHedySessions_(targetUser, 20);
  const todos = syncHedyTodos_(targetUser);
  const highlights = syncHedyHighlights_(targetUser, 50);

  logEvent(targetUser, 'Success', `Hedy API sync completed: sessions +${sessions.created}/${sessions.updated}, todos +${todos.created}/${todos.skipped}, highlights +${highlights.created}/${highlights.updated}.`, {
    sessions: sessions,
    todos: todos,
    highlights: highlights
  });

  return {
    sessions: sessions,
    todos: todos,
    highlights: highlights
  };
}

function syncRecentHedyMeetings(currentUser) {
  const targetUser = String(currentUser || '').trim();
  if (!getActiveTeamMembers_().includes(targetUser)) {
    throw new Error('未指定有效使用者，無法同步 Hedy');
  }

  const props = PropertiesService.getScriptProperties();
  const cursorKey = getHedySyncCursorKey_(targetUser);
  const cursor = String(props.getProperty(cursorKey) || '').trim();
  const mode = cursor ? 'older' : 'latest';
  const syncIndex = createHedySyncIndex_(targetUser);
  const sessions = syncRecentHedySessions_(targetUser, HEDY_HEADER_SYNC_LIMIT, cursor, syncIndex);
  const todos = syncHedyTodosForSessions_(targetUser, sessions.sessionIds, syncIndex);
  const pagination = sessions.pagination || {};
  const nextCursor = String(pagination.next || '').trim();
  const hasMore = Boolean(pagination.hasMore && nextCursor);

  if (nextCursor) {
    props.setProperty(cursorKey, nextCursor);
  } else if (mode === 'older') {
    props.deleteProperty(cursorKey);
  }

  logEvent(targetUser, 'Success', `Hedy API ${mode} sync completed: limit ${HEDY_HEADER_SYNC_LIMIT}, sessions +${sessions.created}/${sessions.updated}, todos +${todos.created}/${todos.skipped} from ${todos.sessions} sessions, highlights ${sessions.highlights}, hasMore ${hasMore}.`, {
    limit: HEDY_HEADER_SYNC_LIMIT,
    mode: mode,
    hasMore: hasMore,
    sessions: sessions,
    todos: todos
  });

  return {
    limit: HEDY_HEADER_SYNC_LIMIT,
    mode: mode,
    hasMore: hasMore,
    sessions: sessions,
    todos: todos
  };
}

function syncLatestHedyMeetings(currentUser) {
  const targetUser = String(currentUser || '').trim();
  if (!getActiveTeamMembers_().includes(targetUser)) {
    throw new Error('未指定有效使用者，無法同步 Hedy');
  }

  const syncIndex = createHedySyncIndex_(targetUser);
  const sessions = syncRecentHedySessions_(targetUser, HEDY_HEADER_SYNC_LIMIT, '', syncIndex);
  const todos = syncHedyTodosForSessions_(targetUser, sessions.sessionIds, syncIndex);

  logEvent(targetUser, 'Success', `Hedy API scheduled latest sync completed: limit ${HEDY_HEADER_SYNC_LIMIT}, sessions +${sessions.created}/${sessions.updated}, todos +${todos.created}/${todos.skipped} from ${todos.sessions} sessions, highlights ${sessions.highlights}.`, {
    limit: HEDY_HEADER_SYNC_LIMIT,
    mode: 'latest',
    sessions: sessions,
    todos: todos
  });

  return {
    limit: HEDY_HEADER_SYNC_LIMIT,
    mode: 'latest',
    sessions: sessions,
    todos: todos
  };
}

function getScheduledHedySyncUser_() {
  const configuredUser = String(PropertiesService.getScriptProperties().getProperty('HEDY_SYNC_USER') || '').trim();
  const members = getActiveTeamMembers_();
  if (members.includes(configuredUser)) {
    return configuredUser;
  }
  return members[0];
}

function scheduledSyncHedyForConfiguredUser() {
  return syncLatestHedyMeetings(getScheduledHedySyncUser_());
}

function testHedyApiConnection(apiKey) {
  hedyApiRequest_('/sessions?limit=1', { apiKey: apiKey });
  return true;
}

function authorizeHedyApiAccess() {
  const response = UrlFetchApp.fetch(HEDY_API_BASE_URL + '/docs', {
    method: 'get',
    muteHttpExceptions: true
  });

  return `Hedy API access checked. HTTP ${response.getResponseCode()}`;
}

function listHedyWebhooks() {
  const response = hedyApiRequest_('/webhooks');
  return unwrapHedyList_(response);
}

function createHedyWebhookForUser(currentUser) {
  const targetUser = String(currentUser || '').trim();
  if (!getActiveTeamMembers_().includes(targetUser)) {
    throw new Error('未指定有效使用者，無法建立 Hedy webhook');
  }

  const config = getWebhookConfig();
  if (!config.url || config.url.indexOf('DEPLOY_WEB_APP_FIRST') >= 0) {
    throw new Error('尚未部署 Web App，無法建立 Hedy webhook');
  }

  const webhookUrl = config.url + '?token=' + encodeURIComponent(config.token) + '&user=' + encodeURIComponent(targetUser);
  const response = hedyApiRequest_('/webhooks', {
    method: 'post',
    payload: {
      url: webhookUrl,
      events: HEDY_WEBHOOK_EVENTS
    }
  });
  const webhook = response && response.data ? response.data : response;

  if (webhook && webhook.signingSecret) {
    PropertiesService.getScriptProperties().setProperty(`HEDY_WEBHOOK_SIGNING_SECRET_${targetUser}`, webhook.signingSecret);
  }

  logEvent(targetUser, 'Success', 'Created Hedy API webhook subscription.', {
    webhookId: webhook && webhook.id,
    events: HEDY_WEBHOOK_EVENTS
  });

  return webhook;
}

/**
 * ==========================================
 * Webhook Endpoint (供 Hedy AI 呼叫)
 * ==========================================
 */
function doPost(e) {
  try {
    const params = e && e.parameter ? e.parameter : {};
    const token = params.token;
    const targetUser = resolveWebhookUser_(params);
    
    const props = PropertiesService.getScriptProperties();
    const validToken = props.getProperty('WEBHOOK_TOKEN') || 'hedy-ai-secret-token-123';
    
    if (token !== validToken) {
      logEvent(targetUser, 'Error', 'Unauthorized Webhook Call', { providedToken: token });
      return createJsonResponse({ status: 'error', message: 'Invalid token' }, 401);
    }

    if (!targetUser) {
      return createJsonResponse({ status: 'error', message: '未指定或無效的使用者。請在 webhook URL 傳入有效的 user 參數。' }, 400);
    }

    if (!e.postData || !e.postData.contents) {
      throw new Error("Empty payload");
    }
    const rawPayload = JSON.parse(e.postData.contents);
    const webhookEvent = normalizeHedyWebhookEvent_(rawPayload);
    const payload = webhookEvent.data;
    const eventType = webhookEvent.eventType;

    // 依照事件類型進行分流處理
    switch (eventType) {
      case 'session.created':
        logEvent(targetUser, 'Success', `Received ${eventType}, ignored until session.ended.`, webhookEvent.rawPayload);
        return createJsonResponse({
          status: 'success',
          message: `Ignored ${eventType}; waiting for session.ended.`
        }, 200);

      case 'session.ended':
      case 'session.exported':
        const meetingRecord = createMeetingRecord_(payload, eventType, 'Hedy AI', webhookEvent.rawPayload);
        const meetingResult = upsertMeetingRecord_(targetUser, meetingRecord);
        const linkedTaskCount = linkExistingTasksToMeeting_(targetUser, meetingResult.meeting);
        const highlightCount = upsertEmbeddedHighlights_(targetUser, payload, meetingResult.meeting, webhookEvent.rawPayload);
        logEvent(
          targetUser,
          'Success',
          `Received ${eventType}, ${meetingResult.created ? 'created' : 'updated'} meeting record, linked ${linkedTaskCount} tasks and synced ${highlightCount} highlights.`,
          webhookEvent.rawPayload
        );
        return createJsonResponse({
          status: 'success',
          message: `Processed ${eventType}, ${meetingResult.created ? 'created' : 'updated'} meeting record.`
        }, 200);

      case 'todo.exported':
        const todoResult = upsertTodoExportTask_(payload, targetUser);
        logEvent(
          targetUser,
          'Success',
          todoResult.created
            ? `Received ${eventType}, created task ${todoResult.taskId}.`
            : `Received ${eventType}, skipped duplicate task ${todoResult.taskId}.`,
          webhookEvent.rawPayload
        );
        return createJsonResponse({
          status: 'success',
          message: todoResult.created
            ? `Processed ${eventType}, created 1 task.`
            : `Processed ${eventType}, duplicate skipped.`
        }, 200);

      case 'highlight.created':
        const highlightRecord = createHighlightRecord_(payload, targetUser, webhookEvent.rawPayload);
        const highlightResult = upsertHighlightRecord_(targetUser, highlightRecord);
        logEvent(
          targetUser,
          'Success',
          `Received ${eventType}, ${highlightResult.created ? 'created' : 'updated'} highlight ${highlightResult.highlight.highlightId}.`,
          webhookEvent.rawPayload
        );
        return createJsonResponse({
          status: 'success',
          message: `Processed ${eventType}, ${highlightResult.created ? 'created' : 'updated'} highlight.`
        }, 200);

      default:
        const fallbackMeetingRecord = createMeetingRecord_(payload, eventType, 'Hedy AI', webhookEvent.rawPayload);
        upsertMeetingRecord_(targetUser, fallbackMeetingRecord);
        logEvent(targetUser, 'Success', `Received ${eventType}, stored meeting record only.`, webhookEvent.rawPayload);
        return createJsonResponse({ status: 'success', message: `Processed ${eventType}, stored meeting record.` }, 200);
    }

  } catch (err) {
    const members = getRegisteredTeamMembers_();
    const fallbackUser = e && e.parameter ? (resolveWebhookUser_(e.parameter) || members[0]) : members[0];
    logEvent(fallbackUser, 'Error', err.toString(), e && e.postData ? e.postData.contents : 'No data');
    return createJsonResponse({ status: 'error', message: err.toString() }, 500);
  }
}

function createJsonResponse(data, code) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 取得與儲存 Webhook 設定
 */
function getWebhookConfig(currentUser) {
  const props = PropertiesService.getScriptProperties();
  let token = props.getProperty('WEBHOOK_TOKEN');
  if (!token) {
    token = 'hedy-ai-secret-token-123';
    props.setProperty('WEBHOOK_TOKEN', token);
  }
  const hedyApiKey = getHedyApiKey_();
  const rawUrl = ScriptApp.getService().getUrl();
  return {
    url: normalizePublicWebAppUrl_(rawUrl),
    token: token,
    hasHedyApiKey: Boolean(hedyApiKey),
    hedyApiKeyPreview: hedyApiKey ? `${hedyApiKey.slice(0, 6)}...${hedyApiKey.slice(-4)}` : '',
    teamMembers: getTeamMemberAdminData(currentUser)
  };
}

function saveWebhookConfig(newToken, hedyApiKey) {
  PropertiesService.getScriptProperties().setProperty('WEBHOOK_TOKEN', newToken);
  if (hedyApiKey !== undefined && hedyApiKey !== null) {
    const normalizedKey = String(hedyApiKey || '').trim();
    if (normalizedKey) {
      PropertiesService.getScriptProperties().setProperty('HEDY_API_KEY', normalizedKey);
    }
  }
  return true;
}

function normalizePublicWebAppUrl_(url) {
  const raw = String(url || '').trim();
  if (!raw) return '';

  return raw.replace(/^https:\/\/script\.google\.com\/a\/[^/]+\/macros\//, 'https://script.google.com/macros/');
}

function debugTestExtractExportedSession_() {
  const payload = {
    title: '業務成長與AI轉型策略會議',
    recap: '簡短 recap',
    meeting_minutes: '## 行動項目\n- 項目一\n- 項目二',
    todos: [],
    待辦事項: [
      '業務團隊需在本周內提交實現兩億元增長的前五大機會清單，並標註每項的預估收入占比',
      '製作市場-收入-人力矩陣'
    ]
  };

  const items = extractActionItemsFromPayload_(payload);
  if (items.length !== 2) {
    throw new Error('Expected 2 action items from 待辦事項');
  }

  if (items[0].title.indexOf('兩億元增長') === -1) {
    throw new Error('Expected first action item title to contain source text');
  }

  const meeting = buildMeetingDetailFields_(payload);
  if (meeting.summaryMarkdown.indexOf('簡短 recap') === -1) {
    throw new Error('Expected summaryMarkdown to include recap');
  }

  if (meeting.actionItemsMarkdown.indexOf('市場-收入-人力矩陣') === -1) {
    throw new Error('Expected actionItemsMarkdown to include extracted action item');
  }

  return 'ok';
}

function debugTestExtractPlainTextActionSection_() {
  const payload = {
    recap: `本次118分鐘會議深入探討程曦資訊整合股份有限公司的業務成長策略。

待辦事項:
• 業務團隊需在本周內提交實現兩億元增長的前五大機會清單，並標註每項的預估收入占比
• 製作『市場-收入-人力』矩陣，列出公共部門、金融、能源、醫療、長照等主要市場的預估收入、佔比及所需人力與AI產能
• 將AI客服釋出的產能，對應到代管業務的500間套房與移工宿舍擴張計畫，確認AI模型支援可行性
• 點賺部門需在每週營運會上匯報AI服務的報名人數、轉換率與預估收入儀表板，追蹤一千萬目標
• 在共享儀表板上加入『AI產能對各市場收入貢獻』指標，分市場顯示公共部門、長照與金融支付的營收貢獻
• 將AI服務一千萬營收目標拆解為每月、每週的具體數字，並設定『AI服務營收占比』指標，納入本季OKR追蹤

決議事項:
• 明確業務成長來源`
  };

  const items = extractActionItemsFromPayload_(payload);
  if (items.length !== 6) {
    throw new Error(`Expected 6 action items from plain-text 待辦事項 section, got ${items.length}`);
  }

  if (items[5].title.indexOf('AI服務營收占比') === -1) {
    throw new Error('Expected sixth action item to preserve tail content');
  }

  return 'ok';
}
