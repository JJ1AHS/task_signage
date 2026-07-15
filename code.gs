/**
 * タスクサイネージ — サーバー側 (Google Apps Script)
 *
 * 必要な設定:
 *   1. エディタ左メニュー「サービス」から以下の拡張サービスを追加
 *        - Tasks API        (識別子: Tasks)
 *        - People API       (識別子: People)
 *        - Google Chat API  (識別子: Chat) … スペース名の表示に使用
 *      ※カレンダーは標準サービス(CalendarApp)のため追加不要
 *   2. デプロイ > 新しいデプロイ > ウェブアプリ
 *        - 実行ユーザー: 自分
 *        - アクセス: 自分のみ（サイネージPCで自分のアカウントにログインして表示）
 *   3. コード更新時は「デプロイを管理」→ 編集 → 新バージョン で反映（URL不変）
 */

var TZ = Session.getScriptTimeZone() || 'Asia/Tokyo';
var WEEK_DAYS = 7;             // 「今週」= 明日から7日後まで
var CONTACT_CACHE_SEC = 21600; // 連絡先キャッシュ 6時間

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('タスクサイネージ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ================ データ取得 ================ */

function getSignageData() {
  var todayStr = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var limitStr = addDays_(todayStr, WEEK_DAYS);
  var todayStartIso = new Date(
    Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'00:00:00XXX")
  ).toISOString();

  var groups = { today: [], week: [], overdue: [], doneToday: [] };

  fetchAllPages_(function (pt) {
    return Tasks.Tasklists.list({ maxResults: 100, pageToken: pt });
  }).forEach(function (list) {
    fetchAllPages_(function (pt) {
      return Tasks.Tasks.list(list.id, {
        showCompleted: false, showAssigned: true,
        maxResults: 100, pageToken: pt
      });
    }).forEach(function (t) {
      classify_(t, list, todayStr, limitStr, groups);
    });

    fetchAllPages_(function (pt) {
      return Tasks.Tasks.list(list.id, {
        showCompleted: true, showHidden: true, showAssigned: true,
        completedMin: todayStartIso, maxResults: 100, pageToken: pt
      });
    }).forEach(function (t) {
      if (t.status === 'completed') {
        groups.doneToday.push(toItem_(t, list, todayStr));
      }
    });
  });

  groups.overdue.sort(byDue_);
  groups.week.sort(byDue_);
  groups.today.sort(function (a, b) {
    return (a.due ? 0 : 1) - (b.due ? 0 : 1);
  });

  return {
    groups: groups,
    events: getTodayEvents_(),
    updatedAt: Utilities.formatDate(new Date(), TZ, 'M/d(E) HH:mm'),
    settings: getSettings_()
  };
}

function classify_(t, list, todayStr, limitStr, groups) {
  var item = toItem_(t, list, todayStr);
  if (!item.due) {
    item.dueLabel = '期限なし';
    groups.today.push(item);
  } else if (item.due < todayStr) {
    groups.overdue.push(item);
  } else if (item.due === todayStr) {
    groups.today.push(item);
  } else if (item.due <= limitStr) {
    groups.week.push(item);
  }
}

function toItem_(t, list, todayStr) {
  var due = t.due ? t.due.slice(0, 10) : null;
  var item = {
    id: t.id, listId: list.id, listTitle: list.title,
    title: t.title || '(無題)', notes: t.notes || '',
    due: due, dueLabel: due ? fmtDate_(due) : '', overdueDays: 0,
    space: spaceName_(t),
    assigned: !!t.assignmentInfo,
    webViewLink: t.webViewLink || '',
    position: t.position || ''
  };
  if (due && due < todayStr) {
    item.overdueDays = Math.round(
      (new Date(todayStr + 'T00:00:00Z') - new Date(due + 'T00:00:00Z')) / 86400000
    );
  }
  return item;
}

/* ================ カレンダー（当日の予定） ================ */

function getTodayEvents_() {
  var out = [];
  try {
    CalendarApp.getDefaultCalendar().getEventsForDay(new Date()).forEach(function (ev) {
      out.push({
        title: ev.getTitle() || '(無題)',
        allDay: ev.isAllDayEvent(),
        start: Utilities.formatDate(ev.getStartTime(), TZ, 'HH:mm'),
        end: Utilities.formatDate(ev.getEndTime(), TZ, 'HH:mm'),
        location: ev.getLocation() || ''
      });
    });
  } catch (e) { /* カレンダー未承認時は空表示 */ }
  out.sort(function (a, b) {
    if (a.allDay !== b.allDay) return a.allDay ? -1 : 1;
    return a.start.localeCompare(b.start);
  });
  return out;
}

/* ================ 画面設定 ================ */

var SETTINGS_DEFAULT = {
  order: ['clock', 'tasks', 'cal'],
  show: { clock: true, tasks: true, cal: true },
  clockSizeL: 300, clockSizeP: 220,
  calSizeL: 300, calSizeP: 220,
  clockNarrow: { analog: false, digital: true },
  clockWide:   { analog: true,  digital: true },
  analogStyle: 1, digitalStyle: 1,
  taskSort: 'due',
  theme: 'light'
};

function saveSettings(s) {
  PropertiesService.getUserProperties()
    .setProperty('signage_settings_v2', JSON.stringify(s));
  return { ok: true };
}

function getSettings_() {
  var raw = PropertiesService.getUserProperties().getProperty('signage_settings_v2');
  var out = JSON.parse(JSON.stringify(SETTINGS_DEFAULT));
  if (raw) {
    try {
      var s = JSON.parse(raw);
      for (var k in out) if (k in s) out[k] = s[k];
    } catch (e) { /* 既定値のまま */ }
  }
  return out;
}

/* ================ タスク操作 ================ */

function completeTask(listId, taskId) {
  Tasks.Tasks.patch({ status: 'completed' }, listId, taskId);
  return getSignageData();
}

function reopenTask(listId, taskId) {
  Tasks.Tasks.patch({ status: 'needsAction' }, listId, taskId);
  return getSignageData();
}

function snoozeToTomorrow(listId, taskId) {
  var tomorrow = addDays_(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), 1);
  return changeDue(listId, taskId, tomorrow);
}

function changeDue(listId, taskId, dateStr) {
  Tasks.Tasks.patch({ due: dateStr + 'T00:00:00.000Z' }, listId, taskId);
  return getSignageData();
}

function deleteTask(listId, taskId) {
  Tasks.Tasks.remove(listId, taskId);
  return getSignageData();
}

/* ================ 連絡先・依頼メール ================ */

function getContacts() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('contacts_v2');
  if (hit) return JSON.parse(hit);

  var map = {};

  // Workspaceディレクトリ（組織のユーザー）
  try {
    var pt1;
    do {
      var res1 = People.People.listDirectoryPeople({
        readMask: 'names,emailAddresses',
        sources: ['DIRECTORY_SOURCE_TYPE_DOMAIN_PROFILE',
                  'DIRECTORY_SOURCE_TYPE_DOMAIN_CONTACT'],
        pageSize: 200, pageToken: pt1
      });
      (res1.people || []).forEach(function (p) { collect_(p, map); });
      pt1 = res1.nextPageToken;
    } while (pt1);
  } catch (e) { /* Workspace以外・権限なしはスキップ */ }

  // 自分の連絡先
  try {
    var pt2;
    do {
      var res2 = People.People.Connections.list('people/me', {
        personFields: 'names,emailAddresses', pageSize: 200, pageToken: pt2
      });
      (res2.connections || []).forEach(function (p) { collect_(p, map); });
      pt2 = res2.nextPageToken;
    } while (pt2);
  } catch (e) { /* スキップ */ }

  // その他の連絡先（Gmailでやり取りしただけの相手 = Gmailサジェストの元データ）
  try {
    var pt3;
    do {
      var res3 = People.OtherContacts.list({
        readMask: 'names,emailAddresses', pageSize: 500, pageToken: pt3
      });
      (res3.otherContacts || []).forEach(function (p) { collect_(p, map); });
      pt3 = res3.nextPageToken;
    } while (pt3);
  } catch (e) { /* スキップ */ }

  var out = Object.keys(map).map(function (k) { return map[k]; });
  out.sort(function (a, b) { return a.name.localeCompare(b.name, 'ja'); });
  cache.put('contacts_v2', JSON.stringify(out), CONTACT_CACHE_SEC);
  return out;
}

function collect_(p, map) {
  var email = p.emailAddresses && p.emailAddresses[0] && p.emailAddresses[0].value;
  if (!email || map[email.toLowerCase()]) return;
  var name = (p.names && p.names[0] && p.names[0].displayName) || email;
  map[email.toLowerCase()] = { name: name, email: email };
}

/**
 * 依頼メール送信
 * payload: { email, name, taskTitle, taskNotes, requestDue('yyyy-MM-dd'|null), template('std'|'polite'|'short') }
 */
function sendRequestMail(payload) {
  var me = Session.getActiveUser().getEmail();
  var dueLabel = payload.requestDue ? fmtDate_(payload.requestDue) : '特になし';
  var lines;

  if (payload.template === 'polite') {
    lines = [
      payload.name + ' 様',
      '',
      'いつもお世話になっております。',
      'お忙しいところ恐れ入りますが、下記のタスクについてご対応をお願いできますでしょうか。',
      '',
      '■ タスク　： ' + payload.taskTitle,
      '■ 希望期限： ' + dueLabel + '（ご都合が難しい場合はご相談ください）'
    ];
    if (payload.taskNotes) lines.push('■ 詳細　　： ' + payload.taskNotes);
    lines.push('', 'お手数をおかけしますが、何卒よろしくお願い申し上げます。');
  } else if (payload.template === 'short') {
    lines = [
      payload.name + ' さん',
      '',
      'お疲れさまです。下記のタスクをお願いします。',
      '',
      '・タスク: ' + payload.taskTitle,
      '・期限: ' + dueLabel
    ];
    if (payload.taskNotes) lines.push('・詳細: ' + payload.taskNotes);
    lines.push('', 'よろしくお願いします！');
  } else { // std
    lines = [
      payload.name + ' 様',
      '',
      'お世話になっております。',
      '下記のタスクについて、ご対応をお願いいたします。',
      '',
      '■ タスク　： ' + payload.taskTitle,
      '■ 希望期限： ' + dueLabel
    ];
    if (payload.taskNotes) lines.push('■ 詳細　　： ' + payload.taskNotes);
    lines.push('', 'ご不明点があればご連絡ください。', 'よろしくお願いいたします。');
  }

  lines.push('', '---', 'このメールはタスクサイネージから送信されています。', '差出人: ' + me);
  GmailApp.sendEmail(payload.email, '【タスク依頼】' + payload.taskTitle, lines.join('\n'));
  return { ok: true };
}

/* ================ ユーティリティ ================ */

/**
 * Chatスペース由来のタスクならスペース名を返す（6時間キャッシュ）。
 * Chat APIが未追加・権限なしでも "Chatスペース" と表示して動作継続。
 */
function spaceName_(t) {
  var info = t.assignmentInfo && t.assignmentInfo.spaceInfo;
  if (!info || !info.space) return '';
  var cache = CacheService.getScriptCache();
  var key = 'space_' + info.space;
  var hit = cache.get(key);
  if (hit) return hit;
  var name = 'Chatスペース';
  try {
    var sp = Chat.Spaces.get(info.space);
    if (sp && sp.displayName) name = sp.displayName;
  } catch (e) { /* フォールバック表示のまま */ }
  cache.put(key, name, 21600);
  return name;
}

function fetchAllPages_(fetcher) {
  var out = [], pt = null;
  do {
    var res = fetcher(pt) || {};
    out = out.concat(res.items || []);
    pt = res.nextPageToken;
  } while (pt);
  return out;
}

function addDays_(dateStr, n) {
  var d = new Date(dateStr + 'T00:00:00Z');
  d.setUTCDate(d.getUTCDate() + n);
  return d.toISOString().slice(0, 10);
}

function fmtDate_(dateStr) {
  var d = new Date(dateStr + 'T00:00:00Z');
  var youbi = ['日', '月', '火', '水', '木', '金', '土'][d.getUTCDay()];
  return (d.getUTCMonth() + 1) + '/' + d.getUTCDate() + '(' + youbi + ')';
}

function byDue_(a, b) {
  return (a.due || '9999').localeCompare(b.due || '9999');
}
