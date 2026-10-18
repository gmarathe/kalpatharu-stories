/**
 * ಕಲ್ಪತರು ಕಥೆಗಳು — Senja ಬದಲಿ ವ್ಯವಸ್ಥೆ
 * - ಸದಸ್ಯರ ವಿಡಿಯೋ: Google Drive → "Kalpatharu Stories/<form>"
 * - ಎಲ್ಲಾ ಕಥೆಗಳು: Google Sheet "Kalpatharu Stories" → Stories ಟ್ಯಾಬ್
 * - ಅನುಮೋದನೆ: Status ಕಾಲಂನಲ್ಲಿ Approved / Pending / Hidden
 * - ವೆಬ್‌ಸೈಟ್ ಫೀಡ್: ?action=feed  (Approved ಮಾತ್ರ)
 * - YouTube ಬ್ಯಾಕಪ್: ಪ್ರತಿ 10 ನಿಮಿಷಕ್ಕೆ ಒಂದು ಹಂತ (ಸಮಯ ಮಿತಿಯೊಳಗೆ)
 */

var ALLOWED_ORIGINS = ['https://stories.kalpatharu.org', 'https://kalpatharu.org', 'https://www.kalpatharu.org'];
var FORMS = { finisher: 'Finisher', hof: 'Hall of Fame', superhealer: 'Super Healer', torchbearer: 'Torchbearer' };
var COLS = ['Timestamp', 'Form', 'Award', 'Kind', 'Status', 'Rating', 'Name', 'Place', 'Phone', 'Email', 'Text',
            'FileName', 'FileId', 'DriveLink', 'Duration', 'SizeMB', 'YouTubeId', 'Source', 'LegacyVideo', 'Shared', 'Notes'];
var CHUNK = 32 * 1024 * 1024; // 256 KB ಗುಣಕ, 50 MB UrlFetch ಮಿತಿಯೊಳಗೆ

/* ---------------- ಮೊದಲ ಬಾರಿ ಸೆಟಪ್ ---------------- */
function setup() {
  var p = PropertiesService.getScriptProperties();
  var root = getRoot_();
  Object.keys(FORMS).concat(['senja-legacy']).forEach(function (k) { getFolder_(k); });
  var ss = getSheet_().getParent();
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('worker').timeBased().everyMinutes(10).create();
  Logger.log('Folder: ' + root.getUrl());
  Logger.log('Sheet: ' + ss.getUrl());
}

function getRoot_() {
  var p = PropertiesService.getScriptProperties(), id = p.getProperty('ROOT');
  if (id) try { return DriveApp.getFolderById(id); } catch (e) {}
  var f = DriveApp.createFolder('Kalpatharu Stories');
  p.setProperty('ROOT', f.getId());
  return f;
}
function getFolder_(key) {
  var p = PropertiesService.getScriptProperties(), id = p.getProperty('F_' + key);
  if (id) try { return DriveApp.getFolderById(id); } catch (e) {}
  var f = getRoot_().createFolder(key === 'senja-legacy' ? 'Senja (ಹಳೆಯ ವಿಡಿಯೋಗಳು)' : FORMS[key]);
  p.setProperty('F_' + key, f.getId());
  return f;
}
function getSheet_() {
  var p = PropertiesService.getScriptProperties(), id = p.getProperty('SHEET');
  var ss = null;
  if (id) try { ss = SpreadsheetApp.openById(id); } catch (e) {}
  if (!ss) {
    ss = SpreadsheetApp.create('Kalpatharu Stories');
    DriveApp.getFileById(ss.getId()).moveTo(getRoot_());
    p.setProperty('SHEET', ss.getId());
    var sh = ss.getSheets()[0]; sh.setName('Stories');
    sh.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold').setBackground('#1B1F3B').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(['Pending', 'Approved', 'Hidden'], true).build();
    sh.getRange(2, 5, 5000, 1).setDataValidation(rule);
    sh.setColumnWidth(11, 360);
    var help = ss.insertSheet('ಸಹಾಯ');
    help.getRange(1, 1, 6, 1).setValues([
      ['ಹೊಸ ಕಥೆ ಬಂದಾಗ Status = Pending. ವೆಬ್‌ಸೈಟ್‌ನಲ್ಲಿ ತೋರಿಸಲು Approved ಆರಿಸಿ.'],
      ['ತೆಗೆಯಲು Hidden ಆರಿಸಿ. ಸಾಲನ್ನು ಅಳಿಸಬೇಡಿ.'],
      ['ವಿಡಿಯೋ ನೋಡಲು DriveLink ಕ್ಲಿಕ್ ಮಾಡಿ.'],
      ['YouTubeId ತುಂಬಿದ್ದರೆ YouTube ಬ್ಯಾಕಪ್ ಆಗಿದೆ.'],
      ['Kind = feedback ಸಾಲುಗಳು (ಕಡಿಮೆ ರೇಟಿಂಗ್) ಎಂದಿಗೂ ವೆಬ್‌ಸೈಟ್‌ಗೆ ಹೋಗುವುದಿಲ್ಲ.'],
      ['ವೆಬ್‌ಸೈಟ್ 5 ನಿಮಿಷದೊಳಗೆ ಹೊಸ ಬದಲಾವಣೆ ತೋರಿಸುತ್ತದೆ.']
    ]);
  }
  return ss.getSheetByName('Stories');
}

/* ---------------- ವೆಬ್ ಎಂಡ್‌ಪಾಯಿಂಟ್‌ಗಳು ---------------- */
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doGet(e) {
  var a = (e && e.parameter && e.parameter.action) || '';
  if (a === 'feed') return json_({ ok: true, stories: feed_() });
  return json_({ ok: true, service: 'kalpatharu-stories' });
}

function doPost(e) {
  try {
    var b = JSON.parse(e.postData.contents || '{}');
    if (!FORMS[b.form]) throw new Error('form');
    if (b.action === 'start') return json_(start_(b));
    if (b.action === 'submit') return json_(submit_(b));
    throw new Error('action');
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function clean_(s, n) { return String(s || '').replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, n || 500); }

function start_(b) {
  var size = Number(b.size) || 0;
  if (size <= 0 || size > 1.6e9) throw new Error('size');
  var ext = String(b.ext || 'mp4').toLowerCase().replace(/[^a-z0-9]/g, '').slice(0, 5) || 'mp4';
  var mime = /^video\/[\w.+-]+$/.test(b.mime || '') ? b.mime : 'video/mp4';
  var stamp = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyyMMdd-HHmmss');
  var who = clean_(b.name, 40).replace(/[\\\/:*?"<>|]/g, '') || 'member';
  var fileName = b.form + '-' + stamp + '-' + who + '-' + Utilities.getUuid().slice(0, 6) + '.' + ext;
  var res = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&fields=id', {
    method: 'post', contentType: 'application/json; charset=UTF-8', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken(), 'X-Upload-Content-Type': mime, 'X-Upload-Content-Length': String(size), Origin: ALLOWED_ORIGINS[0] },
    payload: JSON.stringify({ name: fileName, parents: [getFolder_(b.form).getId()], mimeType: mime })
  });
  var loc = res.getHeaders()['Location'] || res.getHeaders()['location'];
  if (!loc) throw new Error('drive ' + res.getResponseCode());
  return { ok: true, uploadUrl: loc, fileName: fileName };
}

function submit_(b) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = getSheet_();
    var kind = b.kind === 'feedback' ? 'feedback' : 'video';
    var fileId = clean_(b.fileId, 100), link = '';
    if (kind === 'video') {
      if (!fileId && b.fileName) {
        var it = getFolder_(b.form).getFilesByName(clean_(b.fileName, 200));
        if (it.hasNext()) fileId = it.next().getId();
      }
      if (!fileId) throw new Error('file');
      link = 'https://drive.google.com/file/d/' + fileId + '/view';
    }
    var row = {
      Timestamp: new Date(), Form: b.form, Award: FORMS[b.form], Kind: kind, Status: 'Pending',
      Rating: Number(b.rating) || '', Name: clean_(b.name, 80), Place: clean_(b.place, 80),
      Phone: clean_(b.phone, 20), Email: clean_(b.email, 120), Text: clean_(b.text, 2000),
      FileName: clean_(b.fileName, 200), FileId: fileId, DriveLink: link,
      Duration: Number(b.duration) || '', SizeMB: b.size ? Math.round(Number(b.size) / 1e5) / 10 : '',
      YouTubeId: '', Source: 'form', LegacyVideo: '', Shared: '', Notes: ''
    };
    sh.appendRow(COLS.map(function (c) { return row[c]; }));
    notify_(row);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

function notify_(r) {
  try {
    var to = Session.getEffectiveUser().getEmail();
    var url = getSheet_().getParent().getUrl();
    var subj = (r.Kind === 'feedback' ? '⚠️ ಕಡಿಮೆ ರೇಟಿಂಗ್ ಅಭಿಪ್ರಾಯ' : '🎥 ಹೊಸ ' + r.Award + ' ಕಥೆ') + ' — ' + (r.Name || 'ಸದಸ್ಯ');
    var body = 'ಹೆಸರು: ' + r.Name + '\nಊರು: ' + r.Place + '\nWhatsApp: ' + r.Phone + '\nರೇಟಿಂಗ್: ' + r.Rating +
      (r.Text ? '\n\n' + r.Text : '') + (r.DriveLink ? '\n\nವಿಡಿಯೋ: ' + r.DriveLink : '') +
      '\n\nಅನುಮೋದಿಸಲು Sheet ನಲ್ಲಿ Status = Approved ಮಾಡಿ:\n' + url;
    MailApp.sendEmail(to, subj, body);
  } catch (e) {}
}

/* ---------------- ವೆಬ್‌ಸೈಟ್ ಫೀಡ್ ---------------- */
function feed_() {
  var cache = CacheService.getScriptCache(), hit = cache.get('feed');
  if (hit) return JSON.parse(hit);
  var sh = getSheet_(), v = sh.getDataRange().getValues(), h = v.shift(), ix = {};
  h.forEach(function (c, i) { ix[c] = i; });
  var out = [];
  v.forEach(function (r) {
    if (r[ix.Status] !== 'Approved' || r[ix.Kind] === 'feedback') return;
    var isVid = !!(r[ix.FileId] || r[ix.LegacyVideo] || r[ix.YouTubeId]) && r[ix.Kind] !== 'text';
    var d = r[ix.Timestamp] instanceof Date ? Utilities.formatDate(r[ix.Timestamp], 'Asia/Kolkata', 'yyyy-MM') : String(r[ix.Timestamp]).slice(0, 7);
    var s = { type: isVid ? 'video' : 'text', name: r[ix.Name], place: r[ix.Place], program: r[ix.Award] || '',
      source: r[ix.Source] && r[ix.Source] !== 'form' && r[ix.Source] !== 'senja' ? r[ix.Source] : 'ಕಲ್ಪತರು',
      rating: Number(r[ix.Rating]) || 0, date: d, text: r[ix.Text] };
    if (isVid) {
      if (r[ix.FileId] && r[ix.Shared]) s.drive = r[ix.FileId];
      else if (r[ix.LegacyVideo]) s.mux = r[ix.LegacyVideo];
      if (!s.drive && !s.mux) return;
    }
    out.push(s);
  });
  out.reverse();
  try { cache.put('feed', JSON.stringify(out), 300); } catch (e) {}
  return out;
}

/* ---------------- ಹಿನ್ನೆಲೆ ಕೆಲಸ (ಪ್ರತಿ 10 ನಿಮಿಷ) ---------------- */
function worker() {
  var t0 = Date.now(), budget = 4.5 * 60 * 1000;
  var lock = LockService.getScriptLock(); if (!lock.tryLock(1000)) return;
  try {
    shareApproved_();
    while (Date.now() - t0 < budget) {
      if (!legacyStep_(t0, budget) && !youtubeStep_(t0, budget)) break;
    }
  } finally { lock.releaseLock(); }
}

/* Approved ವಿಡಿಯೋಗಳನ್ನು "ಲಿಂಕ್ ಇರುವವರು ನೋಡಬಹುದು" ಮಾಡಿ; Hidden ಆದರೆ ಮತ್ತೆ ಖಾಸಗಿ */
function shareApproved_() {
  var sh = getSheet_(), v = sh.getDataRange().getValues(), h = v[0], ix = {};
  h.forEach(function (c, i) { ix[c] = i; });
  var changed = false;
  for (var i = 1; i < v.length; i++) {
    var id = v[i][ix.FileId]; if (!id) continue;
    var want = v[i][ix.Status] === 'Approved', has = !!v[i][ix.Shared];
    if (want === has) continue;
    try {
      DriveApp.getFileById(id).setSharing(want ? DriveApp.Access.ANYONE_WITH_LINK : DriveApp.Access.PRIVATE, DriveApp.Permission.VIEW);
      sh.getRange(i + 1, ix.Shared + 1).setValue(want ? 'yes' : '');
      changed = true;
    } catch (e) { sh.getRange(i + 1, ix.Notes + 1).setValue('share: ' + e.message); }
  }
  if (changed) CacheService.getScriptCache().remove('feed');
}

/* ತುಂಡು ತುಂಡಾಗಿ ಒಂದು URL ನಿಂದ resumable ಸೆಷನ್‌ಗೆ ಕಳುಹಿಸುವುದು. ಮುಗಿದರೆ ಉತ್ತರದ JSON, ಇಲ್ಲದಿದ್ದರೆ null */
function pump_(job, readChunk, t0, budget) {
  while (Date.now() - t0 < budget) {
    var end = Math.min(job.offset + CHUNK, job.size) - 1;
    var bytes = readChunk(job.offset, end);
    var r = UrlFetchApp.fetch(job.session, { method: 'put', payload: bytes, muteHttpExceptions: true,
      headers: { 'Content-Range': 'bytes ' + job.offset + '-' + end + '/' + job.size } });
    var code = r.getResponseCode();
    if (code === 308) { var rg = r.getHeaders()['Range'] || r.getHeaders()['range']; job.offset = rg ? Number(rg.split('-')[1]) + 1 : end + 1; continue; }
    if (code === 200 || code === 201) return JSON.parse(r.getContentText() || '{}');
    throw new Error('upload ' + code + ' ' + r.getContentText().slice(0, 200));
  }
  return null;
}
function rangeReader_(url, auth) {
  return function (a, b) {
    var h = { Range: 'bytes=' + a + '-' + b }; if (auth) h.Authorization = 'Bearer ' + ScriptApp.getOAuthToken();
    var r = UrlFetchApp.fetch(url, { headers: h, muteHttpExceptions: true });
    if (r.getResponseCode() !== 206 && r.getResponseCode() !== 200) throw new Error('read ' + r.getResponseCode());
    return r.getBlob().getBytes();
  };
}
function head_(url) {
  var r = UrlFetchApp.fetch(url, { headers: { Range: 'bytes=0-0' }, muteHttpExceptions: true });
  var cr = r.getHeaders()['Content-Range'] || r.getHeaders()['content-range'] || '';
  return Number(cr.split('/')[1]) || 0;
}

/* ಹಳೆಯ Senja ವಿಡಿಯೋ → Drive (ಒಂದೊಂದಾಗಿ) */
function legacyStep_(t0, budget) {
  var p = PropertiesService.getScriptProperties();
  var sh = getSheet_(), v = sh.getDataRange().getValues(), h = v[0], ix = {};
  h.forEach(function (c, i) { ix[c] = i; });
  var job = JSON.parse(p.getProperty('LEGACY_JOB') || 'null');
  if (!job) {
    for (var i = 1; i < v.length; i++) {
      if (v[i][ix.LegacyVideo] && !v[i][ix.FileId] && !/^legacy-fail/.test(v[i][ix.Notes])) {
        var src = 'https://stream.mux.com/' + v[i][ix.LegacyVideo] + '/medium.mp4';
        var size = head_(src); if (!size) { sh.getRange(i + 1, ix.Notes + 1).setValue('legacy-fail: size'); continue; }
        var name = 'senja-' + (v[i][ix.FileName] || (v[i][ix.Name] + '.mp4'));
        var res = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&fields=id', {
          method: 'post', contentType: 'application/json; charset=UTF-8', muteHttpExceptions: true,
          headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken(), 'X-Upload-Content-Type': 'video/mp4', 'X-Upload-Content-Length': String(size) },
          payload: JSON.stringify({ name: name, parents: [getFolder_('senja-legacy').getId()], mimeType: 'video/mp4' }) });
        var loc = res.getHeaders()['Location'] || res.getHeaders()['location'];
        if (!loc) { sh.getRange(i + 1, ix.Notes + 1).setValue('legacy-fail: session ' + res.getResponseCode()); continue; }
        job = { row: i + 1, src: src, size: size, offset: 0, session: loc };
        break;
      }
    }
    if (!job) return false;
  }
  try {
    var done = pump_(job, rangeReader_(job.src, false), t0, budget);
    if (done && done.id) {
      sh.getRange(job.row, ix.FileId + 1).setValue(done.id);
      sh.getRange(job.row, ix.DriveLink + 1).setValue('https://drive.google.com/file/d/' + done.id + '/view');
      sh.getRange(job.row, ix.SizeMB + 1).setValue(Math.round(job.size / 1e5) / 10);
      p.deleteProperty('LEGACY_JOB');
      CacheService.getScriptCache().remove('feed');
    } else p.setProperty('LEGACY_JOB', JSON.stringify(job));
  } catch (e) {
    job.tries = (job.tries || 0) + 1;
    if (job.tries >= 3) { sh.getRange(job.row, ix.Notes + 1).setValue('legacy-fail: ' + e.message.slice(0, 150)); p.deleteProperty('LEGACY_JOB'); }
    else p.setProperty('LEGACY_JOB', JSON.stringify(job));
  }
  return true;
}

/* Drive ವಿಡಿಯೋ → YouTube ಬ್ಯಾಕಪ್ (Approved + ಹಳೆಯವು) */
function youtubeStep_(t0, budget) {
  var p = PropertiesService.getScriptProperties();
  if (p.getProperty('YT_OFF')) return false;
  var day = Utilities.formatDate(new Date(), 'America/Los_Angeles', 'yyyyMMdd'); // YouTube ಕೋಟಾ ದಿನ
  var cnt = JSON.parse(p.getProperty('YT_COUNT') || '{}');
  var sh = getSheet_(), v = sh.getDataRange().getValues(), h = v[0], ix = {};
  h.forEach(function (c, i) { ix[c] = i; });
  var job = JSON.parse(p.getProperty('YT_JOB') || 'null');
  if (!job) {
    if ((cnt[day] || 0) >= 5) return false; // ದಿನಕ್ಕೆ 5 (API ಕೋಟಾ ಮಿತಿ)
    for (var i = 1; i < v.length; i++) {
      var r = v[i];
      if (!r[ix.FileId] || r[ix.YouTubeId] || r[ix.Kind] === 'feedback' || /^yt-fail/.test(r[ix.Notes])) continue;
      if (r[ix.Status] !== 'Approved' && r[ix.Source] !== 'senja') continue;
      var f = DriveApp.getFileById(r[ix.FileId]), size = f.getSize();
      var title = ('ಕಲ್ಪತರು ' + (r[ix.Award] || '') + ' ಕಥೆ — ' + r[ix.Name] + (r[ix.Place] ? ', ' + r[ix.Place] : '')).slice(0, 100);
      var meta = { snippet: { title: title, description: 'ಕಲ್ಪತರು ಸದಸ್ಯರ ಅನುಭವ · kalpatharu.org/tv', categoryId: '22', defaultLanguage: 'kn' },
                   status: { privacyStatus: 'unlisted', selfDeclaredMadeForKids: false } };
      var res = UrlFetchApp.fetch('https://www.googleapis.com/upload/youtube/v3/videos?uploadType=resumable&part=snippet,status', {
        method: 'post', contentType: 'application/json; charset=UTF-8', muteHttpExceptions: true,
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken(), 'X-Upload-Content-Type': f.getMimeType() || 'video/mp4', 'X-Upload-Content-Length': String(size) },
        payload: JSON.stringify(meta) });
      var loc = res.getHeaders()['Location'] || res.getHeaders()['location'];
      if (!loc) {
        var txt = res.getContentText();
        if (/quota/i.test(txt)) { cnt[day] = 99; p.setProperty('YT_COUNT', JSON.stringify(cnt)); return false; }
        sh.getRange(i + 1, ix.Notes + 1).setValue('yt-fail: ' + res.getResponseCode() + ' ' + txt.slice(0, 120)); continue;
      }
      job = { row: i + 1, fileId: r[ix.FileId], size: size, offset: 0, session: loc };
      var c2 = {}; c2[day] = (cnt[day] || 0) + 1; p.setProperty('YT_COUNT', JSON.stringify(c2));
      break;
    }
    if (!job) return false;
  }
  try {
    var done = pump_(job, rangeReader_('https://www.googleapis.com/drive/v3/files/' + job.fileId + '?alt=media', true), t0, budget);
    if (done && done.id) {
      sh.getRange(job.row, ix.YouTubeId + 1).setValue(done.id);
      if (done.status && done.status.privacyStatus) sh.getRange(job.row, ix.Notes + 1).setValue('YouTube: ' + done.status.privacyStatus);
      p.deleteProperty('YT_JOB');
    } else p.setProperty('YT_JOB', JSON.stringify(job));
  } catch (e) {
    job.tries = (job.tries || 0) + 1;
    if (job.tries >= 3) { sh.getRange(job.row, ix.Notes + 1).setValue('yt-fail: ' + e.message.slice(0, 150)); p.deleteProperty('YT_JOB'); }
    else p.setProperty('YT_JOB', JSON.stringify(job));
  }
  return true;
}

/* ---------------- Senja ಡೇಟಾ ಆಮದು (ಒಮ್ಮೆ ಮಾತ್ರ) ---------------- */
function importSenja() {
  var sh = getSheet_();
  if (sh.getLastRow() > 1 && PropertiesService.getScriptProperties().getProperty('SENJA_IMPORTED')) { Logger.log('ಈಗಾಗಲೇ ಆಮದು ಆಗಿದೆ'); return; }
  var seed = JSON.parse(UrlFetchApp.fetch('https://raw.githubusercontent.com/gmarathe/kalpatharu-stories/main/data/senja-seed.json').getContentText());
  var rows = seed.map(function (s) {
    var r = { Timestamp: s.date ? s.date + '-01' : '', Form: s.form || '', Award: s.program || '', Kind: s.type === 'video' ? 'video' : 'text',
      Status: 'Approved', Rating: s.rating || '', Name: s.name, Place: s.place || s.role || '', Phone: '', Email: '', Text: s.text || '',
      FileName: s.file || '', FileId: '', DriveLink: '', Duration: '', SizeMB: '', YouTubeId: '',
      Source: s.type === 'video' ? 'senja' : (s.source || 'senja'), LegacyVideo: s.mux || '', Shared: '', Notes: '' };
    return COLS.map(function (c) { return r[c]; });
  });
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, COLS.length).setValues(rows);
  PropertiesService.getScriptProperties().setProperty('SENJA_IMPORTED', '1');
  CacheService.getScriptCache().remove('feed');
  Logger.log('ಆಮದು: ' + rows.length);
}

/* ಪ್ರಗತಿ ನೋಡಲು: Run → status */
function status() {
  var v = getSheet_().getDataRange().getValues(), h = v.shift(), ix = {};
  h.forEach(function (c, i) { ix[c] = i; });
  var legacy = v.filter(function (r) { return r[ix.LegacyVideo]; });
  Logger.log('Senja ವಿಡಿಯೋ Drive ಗೆ: ' + legacy.filter(function (r) { return r[ix.FileId]; }).length + ' / ' + legacy.length);
  Logger.log('YouTube ಬ್ಯಾಕಪ್: ' + v.filter(function (r) { return r[ix.YouTubeId]; }).length);
  Logger.log('Pending: ' + v.filter(function (r) { return r[ix.Status] === 'Pending'; }).length);
}
