/**
 * 재고관리 시스템 - Apps Script 백엔드
 *
 * [설치 방법]
 *  1. 구글시트 열기 → 확장 프로그램 → Apps Script
 *  2. 이 코드 전체를 붙여넣고 저장 (프로젝트 이름은 자유)
 *  3. "배포" → "새 배포" → 유형: 웹 앱
 *     - 다음 사용자 인증 정보로 실행: 나
 *     - 액세스 권한이 있는 사용자: 누구나 (또는 조직 내 사용자)
 *  4. 발급된 웹 앱 URL을 복사 → 대시보드 HTML 의 [설정] 탭에 붙여넣기
 *
 * 주의:
 *  - 시트 구조: 1행=안내문, 2행=헤더, 3행부터=데이터. 이 구조를 유지해야 함.
 *  - 로그(Production/StockLog)는 대시보드에서만 추가/취소. 직접 편집 금지.
 */

const SHEETS = {
  M: 'Materials', P: 'Products', B: 'BOM',
  PL: 'ProductionLog', SL: 'StockLog', S: 'Settings'
};
const TZ = Session.getScriptTimeZone() || 'Asia/Seoul';

function doGet(e)  { return handle_(e); }
function doPost(e) { return handle_(e); }

function handle_(e) {
  try {
    var body = {};
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      body.action = e.parameter.action;
      body.payload = e.parameter.payload ? JSON.parse(e.parameter.payload) : {};
    }
    var action = body.action || 'getAll';
    var payload = body.payload || {};
    var data;
    switch (action) {
      case 'getAll':         data = getAll_(); break;
      case 'addMaterial':    data = addMaterial_(payload); break;
      case 'updateMaterial': data = updateMaterial_(payload); break;
      case 'deleteMaterial': data = deleteMaterial_(payload); break;
      case 'addProduct':     data = addProduct_(payload); break;
      case 'updateProduct':  data = updateProduct_(payload); break;
      case 'deleteProduct':  data = deleteProduct_(payload); break;
      case 'setBOM':         data = setBOM_(payload); break;
      case 'addProduction':  data = addProduction_(payload); break;
      case 'addStock':       data = addStock_(payload); break;
      case 'undoLog':        data = undoLog_(payload); break;
      case 'updateSetting':  data = updateSetting_(payload); break;
      case 'uploadImage':    data = uploadImage_(payload); break;
      case 'deleteImage':    data = deleteImage_(payload); break;
      default: throw new Error('알 수 없는 action: ' + action);
    }
    return out_({ ok: true, data: data });
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

// 헤더=2행, 데이터=3행부터
function readSheet_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('시트 없음: ' + name);
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2) return { sheet: sh, headers: [], rows: [] };
  var headers = sh.getRange(2, 1, 1, lastCol).getValues()[0]
    .map(function (h) { return String(h || '').trim(); });
  if (lastRow < 3) return { sheet: sh, headers: headers, rows: [] };
  var values = sh.getRange(3, 1, lastRow - 2, lastCol).getValues();
  var rows = values.map(function (r, idx) {
    var o = { __row: idx + 3 };
    for (var i = 0; i < headers.length; i++) o[headers[i]] = r[i];
    return o;
  });
  return { sheet: sh, headers: headers, rows: rows };
}

function rowToArray_(headers, obj) {
  return headers.map(function (h) { return obj[h] === undefined ? '' : obj[h]; });
}

function findRow_(rows, key, val) {
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][key]) === String(val)) return rows[i];
  }
  return null;
}

function nowStr_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
}

function pad_(n, len) {
  var s = String(n);
  while (s.length < len) s = '0' + s;
  return s;
}

function nextLogId_(rows, prefix) {
  var max = 0;
  rows.forEach(function (r) {
    var m = String(r.logId || '').match(new RegExp('^' + prefix + '(\\d+)$'));
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  return prefix + pad_(max + 1, 4);
}

function nextEntityId_(rows, prefix) {
  var max = 0;
  rows.forEach(function (r) {
    var m = String(r.id || '').match(new RegExp('^' + prefix + '(\\d+)$'));
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  return prefix + pad_(max + 1, 3);
}

function stripRow_(r) { delete r.__row; return r; }
function fmtDate_(r) {
  if (r['일시'] instanceof Date) {
    r['일시'] = Utilities.formatDate(r['일시'], TZ, 'yyyy-MM-dd HH:mm:ss');
  }
  return r;
}

function getAll_() {
  var m = readSheet_(SHEETS.M);
  var p = readSheet_(SHEETS.P);
  var b = readSheet_(SHEETS.B);
  var pl = readSheet_(SHEETS.PL);
  var sl = readSheet_(SHEETS.SL);
  var s = readSheet_(SHEETS.S);

  var settings = {};
  s.rows.forEach(function (r) { if (r.key) settings[r.key] = r.value; });

  return {
    materials: m.rows.filter(function (r) { return r.id; }).map(stripRow_),
    products:  p.rows.filter(function (r) { return r.id; }).map(stripRow_),
    bom:       b.rows.filter(function (r) { return r['제품id'] && r['재료id']; }).map(stripRow_),
    productionLog: pl.rows.filter(function (r) { return r.logId; }).map(fmtDate_).map(stripRow_).reverse(),
    stockLog:      sl.rows.filter(function (r) { return r.logId; }).map(fmtDate_).map(stripRow_).reverse(),
    settings: settings,
    serverTime: nowStr_()
  };
}

function addMaterial_(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var t = readSheet_(SHEETS.M);
    var id = p.id || nextEntityId_(t.rows, 'M');
    if (findRow_(t.rows, 'id', id)) throw new Error('재료 id 중복: ' + id);
    var defaultThresh = 10;
    var sset = readSheet_(SHEETS.S);
    var setRow = findRow_(sset.rows, 'key', 'defaultThresholdQty');
    if (setRow && setRow.value !== '' && setRow.value !== null) defaultThresh = Number(setRow.value);
    var newRow = {
      'id': id,
      '재료명': p['재료명'] || p.name || '',
      '단위': p['단위'] || p.unit || '개',
      '현재재고': Number(p['현재재고'] !== undefined ? p['현재재고'] : (p.stock || 0)),
      '임계수량': Number(p['임계수량'] !== undefined ? p['임계수량'] : (p.thresholdQty !== undefined ? p.thresholdQty : defaultThresh)),
      '메모': p['메모'] || p.memo || '',
      '그룹': p['그룹'] || p.group || '',
      '이미지URL': p['이미지URL'] || p.imageUrl || ''
    };
    t.sheet.appendRow(rowToArray_(t.headers, newRow));
    return newRow;
  } finally { lock.releaseLock(); }
}

function updateMaterial_(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var t = readSheet_(SHEETS.M);
    var row = findRow_(t.rows, 'id', p.id);
    if (!row) throw new Error('재료 없음: ' + p.id);
    var updates = p.updates || {};
    Object.keys(updates).forEach(function (k) { row[k] = updates[k]; });
    t.sheet.getRange(row.__row, 1, 1, t.headers.length).setValues([rowToArray_(t.headers, row)]);
    return stripRow_(row);
  } finally { lock.releaseLock(); }
}

function deleteMaterial_(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var t = readSheet_(SHEETS.M);
    var row = findRow_(t.rows, 'id', p.id);
    if (!row) throw new Error('재료 없음: ' + p.id);
    var b = readSheet_(SHEETS.B);
    var used = b.rows.some(function (r) { return r['재료id'] === p.id; });
    if (used && !p.force) throw new Error('이 재료를 사용하는 BOM이 있습니다. 먼저 BOM에서 제거하세요.');
    t.sheet.deleteRow(row.__row);
    return { deleted: p.id };
  } finally { lock.releaseLock(); }
}

function addProduct_(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var t = readSheet_(SHEETS.P);
    var id = p.id || nextEntityId_(t.rows, 'P');
    if (findRow_(t.rows, 'id', id)) throw new Error('제품 id 중복: ' + id);
    var newRow = {
      'id': id,
      '제품명': p['제품명'] || p.name || '',
      '메모': p['메모'] || p.memo || ''
    };
    t.sheet.appendRow(rowToArray_(t.headers, newRow));
    if (p.bom && p.bom.length) {
      setBOM_({ productId: id, items: p.bom });
    }
    return newRow;
  } finally { lock.releaseLock(); }
}

function updateProduct_(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var t = readSheet_(SHEETS.P);
    var row = findRow_(t.rows, 'id', p.id);
    if (!row) throw new Error('제품 없음: ' + p.id);
    var updates = p.updates || {};
    Object.keys(updates).forEach(function (k) { row[k] = updates[k]; });
    t.sheet.getRange(row.__row, 1, 1, t.headers.length).setValues([rowToArray_(t.headers, row)]);
    if (p.bom) setBOM_({ productId: p.id, items: p.bom });
    return stripRow_(row);
  } finally { lock.releaseLock(); }
}

function deleteProduct_(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var t = readSheet_(SHEETS.P);
    var row = findRow_(t.rows, 'id', p.id);
    if (!row) throw new Error('제품 없음: ' + p.id);
    var b = readSheet_(SHEETS.B);
    var toDelete = b.rows.filter(function (r) { return r['제품id'] === p.id; })
      .sort(function (a, c) { return c.__row - a.__row; });
    toDelete.forEach(function (r) { b.sheet.deleteRow(r.__row); });
    t.sheet.deleteRow(row.__row);
    return { deleted: p.id };
  } finally { lock.releaseLock(); }
}

function setBOM_(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var t = readSheet_(SHEETS.B);
    var toDelete = t.rows.filter(function (r) { return r['제품id'] === p.productId; })
      .sort(function (a, c) { return c.__row - a.__row; });
    toDelete.forEach(function (r) { t.sheet.deleteRow(r.__row); });
    (p.items || []).forEach(function (item) {
      var qty = Number(item.qty !== undefined ? item.qty : item['1개당_소요량']);
      if (!qty || qty <= 0) return;
      t.sheet.appendRow([p.productId, item.materialId || item['재료id'], qty]);
    });
    return { productId: p.productId, count: (p.items || []).length };
  } finally { lock.releaseLock(); }
}

function addProduction_(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var qty = Number(p.qty || 0);
    if (qty <= 0) throw new Error('생산수량은 1 이상이어야 합니다.');
    var b = readSheet_(SHEETS.B);
    var bomRows = b.rows.filter(function (r) { return r['제품id'] === p.productId; });
    if (bomRows.length === 0) throw new Error('이 제품의 BOM이 등록되지 않았습니다.');

    var m = readSheet_(SHEETS.M);
    var shortages = [];
    bomRows.forEach(function (r) {
      var mat = findRow_(m.rows, 'id', r['재료id']);
      if (!mat) return;
      var need = Number(r['1개당_소요량']) * qty;
      var have = Number(mat['현재재고']);
      if (have < need) {
        shortages.push({
          id: mat.id, name: mat['재료명'], unit: mat['단위'],
          have: have, need: need, short: need - have
        });
      }
    });
    if (shortages.length && !p.force) {
      throw new Error('재료 부족: ' + shortages.map(function (s) {
        return s.name + ' (' + s.short + s.unit + ' 부족)';
      }).join(', '));
    }

    bomRows.forEach(function (r) {
      var mat = findRow_(m.rows, 'id', r['재료id']);
      if (!mat) return;
      var need = Number(r['1개당_소요량']) * qty;
      var newStock = Number(mat['현재재고']) - need;
      m.sheet.getRange(mat.__row, m.headers.indexOf('현재재고') + 1).setValue(newStock);
      mat['현재재고'] = newStock;
    });

    var pl = readSheet_(SHEETS.PL);
    var logId = nextLogId_(pl.rows, 'L');
    var newLog = {
      'logId': logId,
      '일시': nowStr_(),
      '제품id': p.productId,
      '생산수량': qty,
      '메모': p.memo || '',
      '취소여부': 'FALSE'
    };
    pl.sheet.appendRow(rowToArray_(pl.headers, newLog));
    return { logId: logId, log: newLog, shortages: shortages };
  } finally { lock.releaseLock(); }
}

function addStock_(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var amount = Number(p.amount || 0);
    if (amount === 0) throw new Error('변동수량은 0이 아니어야 합니다.');
    var m = readSheet_(SHEETS.M);
    var mat = findRow_(m.rows, 'id', p.materialId);
    if (!mat) throw new Error('재료 없음: ' + p.materialId);
    var newStock = Number(mat['현재재고']) + amount;
    m.sheet.getRange(mat.__row, m.headers.indexOf('현재재고') + 1).setValue(newStock);

    var sl = readSheet_(SHEETS.SL);
    var logId = nextLogId_(sl.rows, 'S');
    var newLog = {
      'logId': logId,
      '일시': nowStr_(),
      '재료id': p.materialId,
      '변동수량': amount,
      '종류': p.type || (amount > 0 ? '입고' : '조정'),
      '메모': p.memo || '',
      '취소여부': 'FALSE'
    };
    sl.sheet.appendRow(rowToArray_(sl.headers, newLog));
    return { logId: logId, log: newLog, newStock: newStock };
  } finally { lock.releaseLock(); }
}

function undoLog_(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    if (p.kind === 'production') {
      var pl = readSheet_(SHEETS.PL);
      var log = findRow_(pl.rows, 'logId', p.logId);
      if (!log) throw new Error('로그 없음');
      if (String(log['취소여부']).toUpperCase() === 'TRUE') throw new Error('이미 취소된 로그입니다.');
      var b = readSheet_(SHEETS.B);
      var m = readSheet_(SHEETS.M);
      b.rows.filter(function (r) { return r['제품id'] === log['제품id']; }).forEach(function (r) {
        var mat = findRow_(m.rows, 'id', r['재료id']);
        if (!mat) return;
        var add = Number(r['1개당_소요량']) * Number(log['생산수량']);
        var newStock = Number(mat['현재재고']) + add;
        m.sheet.getRange(mat.__row, m.headers.indexOf('현재재고') + 1).setValue(newStock);
        mat['현재재고'] = newStock;
      });
      pl.sheet.getRange(log.__row, pl.headers.indexOf('취소여부') + 1).setValue('TRUE');
      return { undone: p.logId };
    } else if (p.kind === 'stock') {
      var sl = readSheet_(SHEETS.SL);
      var log2 = findRow_(sl.rows, 'logId', p.logId);
      if (!log2) throw new Error('로그 없음');
      if (String(log2['취소여부']).toUpperCase() === 'TRUE') throw new Error('이미 취소된 로그입니다.');
      var m2 = readSheet_(SHEETS.M);
      var mat2 = findRow_(m2.rows, 'id', log2['재료id']);
      if (mat2) {
        var newStock2 = Number(mat2['현재재고']) - Number(log2['변동수량']);
        m2.sheet.getRange(mat2.__row, m2.headers.indexOf('현재재고') + 1).setValue(newStock2);
      }
      sl.sheet.getRange(log2.__row, sl.headers.indexOf('취소여부') + 1).setValue('TRUE');
      return { undone: p.logId };
    } else {
      throw new Error("kind는 'production' 또는 'stock'");
    }
  } finally { lock.releaseLock(); }
}

function updateSetting_(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var t = readSheet_(SHEETS.S);
    var row = findRow_(t.rows, 'key', p.key);
    if (row) {
      t.sheet.getRange(row.__row, t.headers.indexOf('value') + 1).setValue(p.value);
    } else {
      t.sheet.appendRow([p.key, p.value, p.desc || '']);
    }
    return { key: p.key, value: p.value };
  } finally { lock.releaseLock(); }
}

/* ====== 이미지 업로드 ====== */
// 이미지 폴더를 둘 부모 폴더 ID. 비우면 Drive 최상위(기존 동작).
var IMAGE_PARENT_FOLDER_ID = '1qRskO6Ok_faGgmEImYJLRULX4MwaIfRk';

function getOrCreateImageFolder_() {
  var props = PropertiesService.getScriptProperties();
  var fid = props.getProperty('imageFolderId');
  if (fid) {
    try {
      var cached = DriveApp.getFolderById(fid);
      if (!cached.isTrashed()) {
        // 부모 폴더가 지정돼 있고 캐시 폴더가 거기 하위가 아니면 무시(아래에서 재탐색/생성).
        if (!IMAGE_PARENT_FOLDER_ID) return cached;
        var parents = cached.getParents();
        while (parents.hasNext()) {
          if (parents.next().getId() === IMAGE_PARENT_FOLDER_ID) return cached;
        }
      }
    } catch (e) { /* 폴더 없어짐 - 재생성 */ }
  }

  var folder;
  if (IMAGE_PARENT_FOLDER_ID) {
    var parent = DriveApp.getFolderById(IMAGE_PARENT_FOLDER_ID);
    var sub = parent.getFoldersByName('재고관리_이미지');
    folder = sub.hasNext() ? sub.next() : parent.createFolder('재고관리_이미지');
  } else {
    var iter = DriveApp.getFoldersByName('재고관리_이미지');
    folder = iter.hasNext() ? iter.next() : DriveApp.createFolder('재고관리_이미지');
  }
  props.setProperty('imageFolderId', folder.getId());
  return folder;
}

function tryDeleteFromUrl_(url) {
  var m = String(url || '').match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (!m) return;
  try { DriveApp.getFileById(m[1]).setTrashed(true); } catch (e) {}
}

function uploadImage_(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    if (!p.data) throw new Error('data(base64) 필수');
    var mime = p.mime || 'image/jpeg';
    var ext = (mime.split('/')[1] || 'jpg').toLowerCase();
    var folder = getOrCreateImageFolder_();
    var name = (p.materialId || 'img') + '_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMddHHmmss') + '.' + ext;
    var blob = Utilities.newBlob(Utilities.base64Decode(p.data), mime, name);
    var file = folder.createFile(blob);
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) { /* 도메인 제한 환경에서는 무시 */ }
    var fileId = file.getId();
    var url = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w800';

    if (p.materialId) {
      var t = readSheet_(SHEETS.M);
      var row = findRow_(t.rows, 'id', p.materialId);
      if (!row) throw new Error('재료 없음: ' + p.materialId);
      var col = t.headers.indexOf('이미지URL');
      if (col < 0) throw new Error('Materials 시트에 "이미지URL" 컬럼이 없습니다. (H2 셀에 헤더 추가 필요)');
      var oldUrl = String(row['이미지URL'] || '');
      if (oldUrl) tryDeleteFromUrl_(oldUrl);
      t.sheet.getRange(row.__row, col + 1).setValue(url);
    }
    return { url: url, fileId: fileId };
  } finally { lock.releaseLock(); }
}

function deleteImage_(p) {
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var t = readSheet_(SHEETS.M);
    var row = findRow_(t.rows, 'id', p.id);
    if (!row) throw new Error('재료 없음: ' + p.id);
    var col = t.headers.indexOf('이미지URL');
    if (col < 0) return { ok: true };
    var oldUrl = String(row['이미지URL'] || '');
    if (oldUrl) tryDeleteFromUrl_(oldUrl);
    t.sheet.getRange(row.__row, col + 1).setValue('');
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/* ====== 마이그레이션: % 기반 임계치 → 절대수량 ======
 * 사용법: Apps Script 편집기 상단 함수 선택 박스에서 "migrate_v2_threshold"
 * 선택 후 ▶️ 실행. 1회만 실행하면 됨.
 */
function migrate_v2_threshold() {
  var sh = ss_().getSheetByName(SHEETS.M);
  if (!sh) throw new Error('Materials 시트를 찾을 수 없습니다.');
  var lastCol = sh.getLastColumn();
  var headers = sh.getRange(2, 1, 1, lastCol).getValues()[0]
    .map(function(h){ return String(h||'').trim(); });

  var pctCol = headers.indexOf('임계치%') + 1;
  var baseCol = headers.indexOf('기준재고') + 1;
  var qtyCol = headers.indexOf('임계수량') + 1;

  if (qtyCol > 0) {
    SpreadsheetApp.getUi().alert('이미 마이그레이션 완료된 시트입니다. (임계수량 컬럼이 이미 존재)');
    return;
  }
  if (pctCol === 0) {
    SpreadsheetApp.getUi().alert('임계치% 컬럼을 찾을 수 없습니다. 시트 구조를 확인하세요.');
    return;
  }

  // 1) 임계치% 헤더를 임계수량으로 변경
  sh.getRange(2, pctCol).setValue('임계수량');

  // 2) 각 행의 값 변환: 기준재고 × 임계치% / 100 → 임계수량
  var lastRow = sh.getLastRow();
  if (lastRow >= 3) {
    for (var r = 3; r <= lastRow; r++) {
      var pct = Number(sh.getRange(r, pctCol).getValue() || 0);
      var base = baseCol > 0 ? Number(sh.getRange(r, baseCol).getValue() || 0) : 0;
      var qty = base > 0 ? Math.max(1, Math.round(base * pct / 100)) : 10;
      sh.getRange(r, pctCol).setValue(qty);
    }
  }

  // 3) Settings: defaultThresholdPct → defaultThresholdQty
  var ssh = ss_().getSheetByName(SHEETS.S);
  if (ssh) {
    var sLastRow = ssh.getLastRow();
    var found = false;
    for (var r2 = 3; r2 <= sLastRow; r2++) {
      var k = String(ssh.getRange(r2, 1).getValue() || '');
      if (k === 'defaultThresholdPct') {
        ssh.getRange(r2, 1).setValue('defaultThresholdQty');
        ssh.getRange(r2, 2).setValue(10);
        ssh.getRange(r2, 3).setValue('신규 재료 추가 시 기본 임계수량 (개)');
        found = true; break;
      }
    }
    if (!found) {
      ssh.appendRow(['defaultThresholdQty', 10, '신규 재료 추가 시 기본 임계수량 (개)']);
    }
  }

  SpreadsheetApp.getUi().alert('✓ 마이그레이션 완료\n\n변경 사항:\n  - Materials.임계치% → 임계수량\n  - 값 자동 변환 (기준재고 × % / 100)\n  - Settings.defaultThresholdQty 설정\n\n기준재고 컬럼은 더 이상 사용하지 않으므로 그대로 두거나 수동 삭제하셔도 됩니다.');
}

/* ====== 마이그레이션 v3: Materials에 '그룹' 컬럼 추가 ======
 * 사용법: Apps Script 편집기 함수 박스에서 "migrate_v3_group" 선택 후 ▶️ 실행. 1회만.
 * 멱등: 이미 컬럼이 있으면 안내만 표시하고 종료.
 */
function migrate_v3_group() {
  var sh = ss_().getSheetByName(SHEETS.M);
  if (!sh) throw new Error('Materials 시트를 찾을 수 없습니다.');
  var lastCol = sh.getLastColumn();
  var headers = sh.getRange(2, 1, 1, lastCol).getValues()[0]
    .map(function (h) { return String(h || '').trim(); });

  if (headers.indexOf('그룹') >= 0) {
    SpreadsheetApp.getUi().alert('이미 적용된 시트입니다. (그룹 컬럼이 이미 존재)');
    return;
  }

  // '이미지URL' 앞에 삽입하면 시트 가독성이 좋지만, 컬럼 삽입은 데이터 이동을 동반.
  // 단순히 끝(마지막 컬럼 +1)에 헤더만 추가 — readSheet_는 헤더명 기반이라 위치 무관.
  var newColIdx = lastCol + 1;
  sh.getRange(2, newColIdx).setValue('그룹');
  // 1행이 안내문 영역이면 같은 줄도 비워둠 (Materials는 1행 안내 사용 안 할 수 있음)
  SpreadsheetApp.getUi().alert(
    '✓ 마이그레이션 완료\n\nMaterials 시트 마지막 컬럼에 "그룹" 헤더가 추가되었습니다.\n' +
    '값이 비어있는 행은 대시보드에서 재료명으로 자동 그룹화됩니다.\n' +
    '필요하면 시트에서 컬럼을 원하는 위치로 드래그해 옮겨도 됩니다.'
  );
}
