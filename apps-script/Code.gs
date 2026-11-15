/**
 * Covenant Life Calendar and Announcements Updates
 * -------------------------------------------------
 * One Apps Script project that:
 *   1. Builds and formats the Google Sheet (tabs, columns, dropdowns, checkboxes).
 *   2. Serves an easy admin page (index.html) protected by a passcode.
 *   3. Serves a public data feed (URL + ?feed=1) that the church website reads.
 *
 * FIRST TIME: run setupSheet() once, then deploy as a Web app
 * (Execute as: Me, Who has access: Anyone).
 */

// ===================== SETTINGS =====================
const SHEET_ID       = '1UUcsFhQJj3ElTh_wfU0-Lpn0N2P_Bwp-OGED2DbuMYw';
const SHEET_TITLE    = 'Covenant Life Calendar and Announcements Updates';
const WEBSITE_URL    = 'https://covenantlifeopc.com/';
const CHURCH_ADDRESS = '1415 E. Green Bay St., Shawano, WI';
const TZ             = 'America/Chicago';
const FEED_CACHE_SECONDS = 120;   // website sees changes within about 2 minutes

// ===================== SHEET LAYOUT =====================
const TAB = {
  HOWTO:  'How To Use',
  SUNDAY: 'This Sunday',
  EVENTS: 'Events',
  ANN:    'Announcements',
  GROUPS: 'Groups'
};

const HEADERS = {};
HEADERS[TAB.EVENTS] = ['ID', 'Event Name', 'Date', 'Start Time', 'End Time', 'Location',
                       'Description', 'Show on Website', 'Last Updated'];
HEADERS[TAB.ANN]    = ['ID', 'Headline', 'Details', 'Show From', 'Show Until',
                       'Banner at Top', 'Show on Website', 'Last Updated'];
HEADERS[TAB.GROUPS] = ['ID', 'Group Name', 'Day of Week', 'Which Weeks', 'Time', 'Location',
                       'Map Address', 'Description', 'Leader', 'Extra Note', 'Show on Website',
                       'Last Updated'];

const WIDTHS = {};
WIDTHS[TAB.EVENTS] = [80, 260, 210, 110, 110, 200, 380, 130, 160];
WIDTHS[TAB.ANN]    = [80, 280, 400, 190, 190, 120, 130, 160];
WIDTHS[TAB.GROUPS] = [80, 220, 120, 130, 110, 180, 260, 340, 140, 260, 130, 160];

// Maps the admin page's field names to sheet column titles.
const FIELDS = {};
FIELDS[TAB.EVENTS] = { id: 'ID', title: 'Event Name', date: 'Date', start: 'Start Time',
  end: 'End Time', where: 'Location', desc: 'Description', show: 'Show on Website' };
FIELDS[TAB.ANN] = { id: 'ID', headline: 'Headline', details: 'Details', from: 'Show From',
  until: 'Show Until', banner: 'Banner at Top', show: 'Show on Website' };
FIELDS[TAB.GROUPS] = { id: 'ID', name: 'Group Name', day: 'Day of Week', weeks: 'Which Weeks',
  time: 'Time', where: 'Location', map: 'Map Address', desc: 'Description', leader: 'Leader',
  note: 'Extra Note', show: 'Show on Website' };

const DATE_COLS = { 'Date': true, 'Show From': true, 'Show Until': true };
const TIME_COLS = { 'Start Time': true, 'End Time': true, 'Time': true };
const CHECK_COLS = { 'Show on Website': true, 'Banner at Top': true };

const SUNDAY_FIELDS = [
  ['Sermon Title',  'The title of this Sunday\'s sermon'],
  ['Bible Passage', 'Example: John 3:16-21'],
  ['Preacher',      'Leave blank if it is our pastor'],
  ['Special Note',  'Optional. Example: Communion Sunday, potluck after worship']
];

const DAYS  = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const WEEKS = ['Every week', '1st', '2nd', '3rd', '4th', 'Last', '1st & 3rd', '2nd & 4th', '1st & 2nd'];

const BURGUNDY = '#511d21';
const GOLD     = '#b0832f';
const CREAM    = '#fbf6ef';

// ===================== MENU =====================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Church Website')
    .addItem('Open the admin page', 'showAdminLink')
    .addItem('Refresh the website now', 'refreshWebsiteNow')
    .addSeparator()
    .addItem('Set or change the admin passcode', 'setPasscode')
    .addItem('Set up or repair this sheet', 'setupSheet')
    .addToUi();
}

/** Keeps ID and Last Updated filled when someone types right into the sheet. */
function onEdit(e) {
  try {
    const sh = e.range.getSheet();
    const name = sh.getName();
    if (!HEADERS[name] || e.range.getRow() < 2) return;
    const headers = HEADERS[name];
    const idCol = headers.indexOf('ID') + 1;
    const updCol = headers.indexOf('Last Updated') + 1;
    for (let r = e.range.getRow(); r <= e.range.getLastRow(); r++) {
      const rowVals = sh.getRange(r, 2, 1, headers.length - 1).getValues()[0];
      const hasContent = rowVals.some(function (v) { return v !== '' && v !== false; });
      if (!hasContent) continue;
      if (!sh.getRange(r, idCol).getValue()) sh.getRange(r, idCol).setValue(newId_());
      sh.getRange(r, updCol).setValue(new Date());
    }
    CacheService.getScriptCache().remove('feed');
  } catch (err) {
    // Simple triggers must never interrupt the person typing.
  }
}

// ===================== SETUP =====================
function setupSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  ss.rename(SHEET_TITLE);
  ss.setSpreadsheetTimeZone(TZ);

  buildHowTo_(ss);
  buildSunday_(ss);
  buildTable_(ss, TAB.EVENTS, seedEvents_());
  buildTable_(ss, TAB.ANN, seedAnnouncements_());
  buildTable_(ss, TAB.GROUPS, seedGroups_());

  // Remove the blank starter tab if it is still there and empty.
  ['Sheet1', 'Hoja 1'].forEach(function (n) {
    const s = ss.getSheetByName(n);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });

  // Tab order and colors
  [TAB.HOWTO, TAB.SUNDAY, TAB.EVENTS, TAB.ANN, TAB.GROUPS].forEach(function (n, i) {
    const s = ss.getSheetByName(n);
    ss.setActiveSheet(s);
    ss.moveActiveSheet(i + 1);
  });
  ss.getSheetByName(TAB.HOWTO).setTabColor(GOLD);
  ss.getSheetByName(TAB.SUNDAY).setTabColor('#1f3b5e');
  ss.getSheetByName(TAB.EVENTS).setTabColor(BURGUNDY);
  ss.getSheetByName(TAB.ANN).setTabColor('#8f2f3a');
  ss.getSheetByName(TAB.GROUPS).setTabColor('#6b2440');
  ss.setActiveSheet(ss.getSheetByName(TAB.HOWTO));

  // Make sure a passcode exists.
  const props = PropertiesService.getScriptProperties();
  let pass = props.getProperty('ADMIN_PASSCODE');
  let msg = 'The sheet is set up.';
  if (!pass) {
    pass = String(Math.floor(100000 + Math.random() * 900000));
    props.setProperty('ADMIN_PASSCODE', pass);
    msg += '\n\nYour admin passcode is: ' + pass +
           '\n\nWrite it down. You can change it any time from the Church Website menu.';
  }
  CacheService.getScriptCache().remove('feed');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { /* run from the editor */ }
}

function buildHowTo_(ss) {
  const sh = ss.getSheetByName(TAB.HOWTO) || ss.insertSheet(TAB.HOWTO);
  sh.clear();
  const rows = [
    ['How to update the Covenant Life website'],
    [''],
    ['THE EASY WAY (recommended)'],
    ['Open the admin page from the Church Website menu at the top of this sheet, or use the bookmark you were given. Type the passcode, then click a big button: This Sunday, Events, Announcements, or Groups.'],
    [''],
    ['WHAT EACH TAB DOES'],
    ['This Sunday: the sermon title and Bible passage shown on the home page. Change it once a week.'],
    ['Events: special services and events. They disappear from the website by themselves after their date.'],
    ['Announcements: short notices. Pick a Show From and Show Until date and they come and go on their own. Check "Banner at Top" to show one across the top of every page.'],
    ['Groups: Bible studies and other regular meetings. The website works out the next meeting date by itself.'],
    [''],
    ['IF YOU TYPE IN THE SHEET DIRECTLY'],
    ['Add a new item on the next empty row. Double-click a date box to pick a date from a calendar. Use the little arrow to pick a time. Leave the ID and Last Updated columns alone; they fill in by themselves.'],
    ['To hide something without deleting it, uncheck "Show on Website".'],
    [''],
    ['HOW FAST DOES THE WEBSITE CHANGE?'],
    ['Within about 2 minutes. To update it right away, use Church Website > Refresh the website now.'],
    [''],
    ['MADE A MISTAKE?'],
    ['Click File > Version history > See version history to go back to an earlier copy. Nothing you do here can break the website design.'],
    [''],
    ['Questions? Contact Bruce Bowers.']
  ];
  sh.getRange(1, 1, rows.length, 1).setValues(rows).setWrap(true)
    .setFontFamily('Arial').setFontSize(12).setVerticalAlignment('top');
  sh.setColumnWidth(1, 820);
  sh.getRange('A1').setFontSize(20).setFontWeight('bold').setFontColor(BURGUNDY);
  [3, 6, 12, 16, 19].forEach(function (r) {
    sh.getRange(r, 1).setFontWeight('bold').setFontColor('#ffffff').setBackground(BURGUNDY);
  });
  sh.setHiddenGridlines(true);
  try { sh.getRange(1, 1, rows.length, 1).protect().setWarningOnly(true); } catch (e) {}
}

function buildSunday_(ss) {
  let sh = ss.getSheetByName(TAB.SUNDAY);
  const isNew = !sh;
  if (isNew) sh = ss.insertSheet(TAB.SUNDAY);
  sh.getRange(1, 1, 1, 3).setValues([['What it is', 'What to show on the website', 'Help']]);
  const existing = isNew ? [] : sh.getRange(2, 2, SUNDAY_FIELDS.length, 1).getValues();
  const rows = SUNDAY_FIELDS.map(function (f, i) {
    return [f[0], existing[i] ? existing[i][0] : '', f[1]];
  });
  sh.getRange(2, 1, rows.length, 3).setValues(rows);
  styleHeader_(sh, 3);
  sh.setColumnWidth(1, 170); sh.setColumnWidth(2, 460); sh.setColumnWidth(3, 380);
  sh.getRange(2, 1, rows.length, 1).setFontWeight('bold').setBackground(CREAM);
  sh.getRange(2, 2, rows.length, 1).setBackground('#ffffff').setFontSize(13).setWrap(true)
    .setBorder(true, true, true, true, true, true, '#d9c9b8', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(2, 3, rows.length, 1).setFontColor('#7a6a6c').setFontStyle('italic').setWrap(true);
  sh.setRowHeights(2, rows.length, 36);
  sh.getRange(2, 1, rows.length, 3).setVerticalAlignment('middle');
  sh.setFrozenRows(1);
  removeExtraColumns_(sh, 3);
}

function buildTable_(ss, name, seedRows) {
  let sh = ss.getSheetByName(name);
  const isNew = !sh;
  if (isNew) sh = ss.insertSheet(name);
  const headers = HEADERS[name];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader_(sh, headers.length);
  headers.forEach(function (h, i) { sh.setColumnWidth(i + 1, WIDTHS[name][i]); });
  sh.setFrozenRows(1);
  sh.setFrozenColumns(2);
  sh.hideColumns(1); // ID column is for the system

  const maxRows = Math.max(sh.getMaxRows(), 300);
  if (sh.getMaxRows() < maxRows) sh.insertRowsAfter(sh.getMaxRows(), maxRows - sh.getMaxRows());
  const body = function (col) { return sh.getRange(2, col, maxRows - 1, 1); };

  headers.forEach(function (h, i) {
    const col = i + 1;
    if (DATE_COLS[h]) {
      body(col).setDataValidation(SpreadsheetApp.newDataValidation().requireDate()
        .setAllowInvalid(false).setHelpText('Double-click to pick a date.').build())
        .setNumberFormat('ddd, mmm d, yyyy');
    } else if (TIME_COLS[h]) {
      body(col).setNumberFormat('@').setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInList(timeList_(), true).setAllowInvalid(false).build());
    } else if (CHECK_COLS[h]) {
      body(col).insertCheckboxes().setHorizontalAlignment('center');
    } else if (h === 'Day of Week') {
      body(col).setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInList(DAYS, true).setAllowInvalid(false).build());
    } else if (h === 'Which Weeks') {
      body(col).setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInList(WEEKS, true).setAllowInvalid(false).build());
    } else if (h === 'Last Updated') {
      body(col).setNumberFormat('mmm d, yyyy h:mm am/pm').setFontColor('#8a7a7c');
    } else if (h === 'Description' || h === 'Details' || h === 'Extra Note') {
      body(col).setWrap(true);
    }
  });
  sh.getRange(2, 1, maxRows - 1, headers.length).setVerticalAlignment('top').setFontSize(11);

  // Soft alternating row colors
  sh.getBandings().forEach(function (b) { b.remove(); });
  sh.getRange(1, 1, maxRows, headers.length)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
    .setHeaderRowColor(BURGUNDY).setFirstRowColor('#ffffff').setSecondRowColor(CREAM);

  // Seed starter rows only on a brand-new tab
  if (isNew && seedRows && seedRows.length) {
    const out = seedRows.map(function (obj) { return objToRow_(name, obj, null); });
    sh.getRange(2, 1, out.length, headers.length).setValues(out);
  }
  removeExtraColumns_(sh, headers.length);
}

function styleHeader_(sh, n) {
  sh.getRange(1, 1, 1, n).setBackground(BURGUNDY).setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(12).setVerticalAlignment('middle').setWrap(true);
  sh.setRowHeight(1, 40);
}

function removeExtraColumns_(sh, keep) {
  const max = sh.getMaxColumns();
  if (max > keep) sh.deleteColumns(keep + 1, max - keep);
}

function timeList_() {
  const out = [];
  for (let m = 6 * 60; m <= 21 * 60 + 45; m += 15) {
    const h24 = Math.floor(m / 60), min = m % 60;
    const h = h24 % 12 === 0 ? 12 : h24 % 12;
    out.push(h + ':' + (min < 10 ? '0' : '') + min + ' ' + (h24 < 12 ? 'AM' : 'PM'));
  }
  return out;
}

// Starter content copied from the current website
function seedEvents_() {
  return [
    { title: 'Community Christmas Carol Sing', date: '2026-12-06', start: '1:00 PM', end: '2:30 PM',
      where: 'Covenant Life', desc: 'Bring the whole family and your neighbors for an afternoon of favorite carols. Everyone is welcome.', show: true }
  ];
}
function seedAnnouncements_() {
  return [
    { headline: 'Community Christmas Carol Sing, Sunday, December 6 at 1:00 PM',
      details: 'Bring a friend! Everyone is welcome.', from: '', until: '2026-12-06', banner: true, show: true }
  ];
}
function seedGroups_() {
  return [
    { name: "Men's Bible Study", day: 'Wednesday', weeks: '2nd & 4th', time: '11:30 AM', where: 'Beans and Books',
      map: 'Beans and Books, Shawano, WI', desc: 'Men of all ages working through a book of the Bible over lunch. Bring your questions.', leader: '', note: '', show: true },
    { name: "Ladies' Bible Study", day: 'Monday', weeks: '1st & 2nd', time: '3:45 PM', where: 'Covenant Life',
      map: CHURCH_ADDRESS, desc: 'Women studying Scripture together, with time to share and pray for one another.', leader: '', note: '', show: true },
    { name: 'Book Club', day: 'Thursday', weeks: '1st', time: '1:30 PM', where: 'Beans and Books',
      map: 'Beans and Books, Shawano, WI', desc: "Good books and good conversation over coffee. A great first step if you're new.", leader: '', note: '', show: true }
  ];
}

// ===================== PASSCODE =====================
function setPasscode() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('Admin passcode',
    'Type a new passcode for the admin page (at least 4 characters).', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const p = res.getResponseText().trim();
  if (p.length < 4) { ui.alert('Please use at least 4 characters.'); return; }
  PropertiesService.getScriptProperties().setProperty('ADMIN_PASSCODE', p);
  ui.alert('Saved. Anyone using the admin page will need the new passcode.');
}

function checkPass_(pass) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('fails') || 0);
  if (fails >= 8) throw new Error('Too many wrong passcodes. Please wait 10 minutes and try again.');
  const real = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSCODE');
  if (!real || String(pass || '').trim() !== real) {
    cache.put('fails', String(fails + 1), 600);
    throw new Error('That passcode is not right. Please try again.');
  }
}

// ===================== WEB APP =====================
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.feed !== undefined) return feedResponse_(p.callback);
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Covenant Life Website Updates')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function feedResponse_(callback) {
  const json = getFeedJson_();
  if (callback && /^[A-Za-z_$][\w$]{0,40}$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function getFeedJson_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('feed');
  if (hit) return hit;
  const json = JSON.stringify(buildFeed_());
  try { cache.put('feed', json, FEED_CACHE_SECONDS); } catch (e) {}
  return json;
}

function buildFeed_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const today = todayStr_();

  const events = readTable_(ss, TAB.EVENTS)
    .filter(function (r) { return r.show && r.title && r.date && r.date >= today; })
    .sort(function (a, b) { return (a.date + toMin_(a.start)).localeCompare(b.date + toMin_(b.start)); })
    .map(function (r) {
      return { title: r.title, date: r.date, time: r.start || '', endTime: r.end || '',
        mins: durationMins_(r.start, r.end), where: r.where || 'Covenant Life',
        address: whereToAddress_(r.where), desc: r.desc || '' };
    });

  const announcements = readTable_(ss, TAB.ANN)
    .filter(function (r) {
      return r.show && r.headline && (!r.from || r.from <= today) && (!r.until || r.until >= today);
    })
    .map(function (r) {
      return { headline: r.headline, details: r.details || '', banner: !!r.banner, until: r.until || '' };
    });

  const groups = readTable_(ss, TAB.GROUPS)
    .filter(function (r) { return r.show && r.name && r.day; })
    .map(function (r) {
      return { name: r.name, wd: DAYS.indexOf(r.day), weeks: weeksToArray_(r.weeks),
        when: whenLabel_(r.day, r.weeks), time: r.time || '', where: r.where || 'Covenant Life',
        map: r.map || whereToAddress_(r.where), desc: r.desc || '', leader: r.leader || '', note: r.note || '' };
    });

  return { updated: new Date().toISOString(), sunday: readSunday_(ss),
           events: events, announcements: announcements, groups: groups };
}

// ===================== ADMIN API (called from index.html) =====================
function adminLogin(pass) {
  checkPass_(pass);
  return { ok: true, websiteUrl: WEBSITE_URL, sheetUrl: SpreadsheetApp.openById(SHEET_ID).getUrl() };
}

function adminGetAll(pass) {
  checkPass_(pass);
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return {
    today: todayStr_(),
    sunday: readSunday_(ss),
    events: readTable_(ss, TAB.EVENTS).filter(function (r) { return r.title; }),
    announcements: readTable_(ss, TAB.ANN).filter(function (r) { return r.headline; }),
    groups: readTable_(ss, TAB.GROUPS).filter(function (r) { return r.name; }),
    options: { times: timeList_(), days: DAYS, weeks: WEEKS }
  };
}

function adminSaveSunday(pass, data) {
  checkPass_(pass);
  return withLock_(function () {
    const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TAB.SUNDAY);
    const vals = [[data.sermonTitle || ''], [data.passage || ''], [data.preacher || ''], [data.note || '']];
    sh.getRange(2, 2, 4, 1).setValues(vals.map(function (v) { return [clean_(v[0])]; }));
    CacheService.getScriptCache().remove('feed');
    return { ok: true };
  });
}

/** type is 'events', 'announcements' or 'groups' */
function adminSaveItem(pass, type, obj) {
  checkPass_(pass);
  const name = typeToTab_(type);
  validateItem_(name, obj);
  return withLock_(function () {
    const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
    const headers = HEADERS[name];
    let row = obj.id ? findRowById_(sh, obj.id) : 0;
    if (!obj.id) obj.id = newId_();
    const values = objToRow_(name, obj, new Date());
    if (!row) row = firstEmptyRow_(sh);
    sh.getRange(row, 1, 1, headers.length).setValues([values]);
    CacheService.getScriptCache().remove('feed');
    return { ok: true, id: obj.id };
  });
}

function adminSetShow(pass, type, id, show) {
  checkPass_(pass);
  const name = typeToTab_(type);
  return withLock_(function () {
    const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
    const row = findRowById_(sh, id);
    if (!row) throw new Error('Could not find that item. Please refresh the page.');
    const headers = HEADERS[name];
    sh.getRange(row, headers.indexOf('Show on Website') + 1).setValue(!!show);
    sh.getRange(row, headers.indexOf('Last Updated') + 1).setValue(new Date());
    CacheService.getScriptCache().remove('feed');
    return { ok: true };
  });
}

function adminDeleteItem(pass, type, id) {
  checkPass_(pass);
  const name = typeToTab_(type);
  return withLock_(function () {
    const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
    const row = findRowById_(sh, id);
    if (!row) throw new Error('Could not find that item. Please refresh the page.');
    sh.deleteRow(row);
    CacheService.getScriptCache().remove('feed');
    return { ok: true };
  });
}

function adminRefreshWebsite(pass) {
  checkPass_(pass);
  CacheService.getScriptCache().remove('feed');
  return { ok: true };
}

// ===================== MENU HELPERS =====================
function refreshWebsiteNow() {
  CacheService.getScriptCache().remove('feed');
  SpreadsheetApp.getUi().alert('Done. The website will show the latest changes the next time the page is loaded.');
}

function showAdminLink() {
  const url = ScriptApp.getService().getUrl();
  const html = url
    ? '<div style="font-family:Arial;font-size:15px;line-height:1.5">' +
      '<p><a href="' + url + '" target="_blank" style="display:inline-block;background:#511d21;color:#fff;' +
      'padding:12px 18px;border-radius:8px;text-decoration:none;font-weight:bold">Open the admin page</a></p>' +
      '<p>Bookmark that page so you can get back to it easily.</p></div>'
    : '<p style="font-family:Arial">The admin page has not been published yet. In the Apps Script editor click Deploy &gt; New deployment &gt; Web app.</p>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(380).setHeight(160), 'Admin page');
}

// ===================== DATA HELPERS =====================
function readTable_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const headers = HEADERS[name];
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, headers.length).getValues();
  const map = FIELDS[name];
  const keys = Object.keys(map);
  return vals.map(function (row) {
    const o = {};
    keys.forEach(function (k) {
      const h = map[k];
      const v = row[headers.indexOf(h)];
      if (DATE_COLS[h]) o[k] = dateStr_(v);
      else if (TIME_COLS[h]) o[k] = timeStr_(v);
      else if (CHECK_COLS[h]) o[k] = v === true;
      else o[k] = v === null || v === undefined ? '' : String(v).trim();
    });
    return o;
  }).filter(function (o) { return o.id || Object.keys(o).some(function (k) { return o[k] && o[k] !== true; }); });
}

function readSunday_(ss) {
  const sh = ss.getSheetByName(TAB.SUNDAY);
  if (!sh) return { sermonTitle: '', passage: '', preacher: '', note: '' };
  const v = sh.getRange(2, 2, 4, 1).getValues();
  return { sermonTitle: String(v[0][0] || '').trim(), passage: String(v[1][0] || '').trim(),
           preacher: String(v[2][0] || '').trim(), note: String(v[3][0] || '').trim() };
}

function objToRow_(name, obj, updated) {
  const headers = HEADERS[name];
  const map = FIELDS[name];
  const row = headers.map(function () { return ''; });
  Object.keys(map).forEach(function (k) {
    const h = map[k];
    const i = headers.indexOf(h);
    let v = obj[k];
    if (k === 'id') v = obj.id || newId_();
    if (DATE_COLS[h]) v = toSheetDate_(v);
    else if (CHECK_COLS[h]) v = !!v;
    else v = clean_(v);
    row[i] = v;
  });
  row[headers.indexOf('Last Updated')] = updated || new Date();
  return row;
}

function validateItem_(name, o) {
  if (name === TAB.EVENTS) {
    if (!clean_(o.title)) throw new Error('Please type a name for the event.');
    if (!o.date) throw new Error('Please pick a date for the event.');
  } else if (name === TAB.ANN) {
    if (!clean_(o.headline)) throw new Error('Please type a headline for the announcement.');
    if (o.from && o.until && o.from > o.until) throw new Error('The "Show until" date is before the "Show from" date.');
  } else if (name === TAB.GROUPS) {
    if (!clean_(o.name)) throw new Error('Please type a name for the group.');
    if (DAYS.indexOf(o.day) < 0) throw new Error('Please pick the day the group meets.');
    if (WEEKS.indexOf(o.weeks) < 0) throw new Error('Please pick which weeks the group meets.');
  }
}

function typeToTab_(type) {
  const t = { events: TAB.EVENTS, announcements: TAB.ANN, groups: TAB.GROUPS }[type];
  if (!t) throw new Error('Unknown section.');
  return t;
}

function findRowById_(sh, id) {
  if (!id || sh.getLastRow() < 2) return 0;
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}

function firstEmptyRow_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return 2;
  const vals = sh.getRange(2, 2, last - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) if (vals[i][0] === '') return i + 2;
  if (last >= sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 50);
  return last + 1;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('Someone else is saving right now. Please try again in a moment.');
  try { return fn(); } finally { lock.releaseLock(); }
}

function newId_() { return Utilities.getUuid().slice(0, 8); }
function clean_(v) { return v === null || v === undefined ? '' : String(v).trim().slice(0, 2000); }
function todayStr_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }

function dateStr_(v) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(v.trim())) return v.trim();
  if (v) { const d = new Date(v); if (!isNaN(d)) return Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); }
  return '';
}

function toSheetDate_(s) {
  if (!s) return '';
  const m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return '';
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0); // noon avoids time-zone slips
}

function timeStr_(v) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, TZ, 'h:mm a');
  return v ? String(v).trim().toUpperCase().replace(/\s+/g, ' ') : '';
}

function toMin_(t) {
  const m = String(t || '').match(/(\d{1,2}):(\d{2})\s*(AM|PM)/i);
  if (!m) return '9999';
  let h = Number(m[1]) % 12; if (/pm/i.test(m[3])) h += 12;
  const total = h * 60 + Number(m[2]);
  return ('0000' + total).slice(-4);
}

function durationMins_(start, end) {
  const a = Number(toMin_(start)), b = Number(toMin_(end));
  if (a === 9999 || b === 9999 || b <= a) return 60;
  return b - a;
}

function whereToAddress_(where) {
  const w = String(where || '').trim();
  if (!w || /covenant life/i.test(w)) return CHURCH_ADDRESS;
  return /shawano/i.test(w) ? w : w + ', Shawano, WI';
}

function weeksToArray_(w) {
  switch (w) {
    case 'Every week': return [1, 2, 3, 4, 5];
    case 'Last': return [-1];
    case '1st': return [1];
    case '2nd': return [2];
    case '3rd': return [3];
    case '4th': return [4];
    case '1st & 3rd': return [1, 3];
    case '2nd & 4th': return [2, 4];
    case '1st & 2nd': return [1, 2];
    default: return [1, 2, 3, 4, 5];
  }
}

function whenLabel_(day, weeks) {
  if (!weeks || weeks === 'Every week') return 'Every ' + day;
  if (weeks === 'Last') return 'Last ' + day + ' of the month';
  return weeks + ' ' + day + (weeks.indexOf('&') >= 0 ? 's' : '');
}
