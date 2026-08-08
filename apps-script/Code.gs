/**
 * 눈에 안 보이는 시간 — 대학생 반복 행동 실태조사
 * 응답 수집용 Google Apps Script
 *
 * 설치 방법은 같은 폴더의 README.md 참고.
 * 이 파일 전체를 Apps Script 편집기의 Code.gs 에 그대로 붙여넣으면 된다.
 */

/** 응답이 쌓일 시트 이름 */
const SHEET_NAME = '응답';

/** doPost 가 실패했을 때 흔적을 남길 시트 이름 */
const ERROR_SHEET = '_오류';

/** 시트를 처음 만들 때 세워둘 기본 열. 나머지 열은 응답이 들어올 때 자동으로 늘어난다. */
const BASE_HEADERS = ['id', 'ts', 'weeklyMin', 'ref'];


/* ══════════════════════════════════════════════════════════════
   진입점
   ══════════════════════════════════════════════════════════════ */

/**
 * 설문 페이지가 보내는 POST 를 받는다.
 *
 * 페이지는 mode:'no-cors' 로 보내기 때문에 브라우저는 이 함수의 응답을
 * 읽지 못한다. 즉 여기서 오류가 나도 응답자 화면에는 "제출 완료"가 뜬다.
 * 그래서 실패를 조용히 넘기지 않고 _오류 시트에 반드시 남긴다.
 */
function doPost(e) {
  const lock = LockService.getScriptLock();

  // 동시에 여러 명이 제출하면 appendRow 가 같은 줄을 덮어쓸 수 있다.
  try {
    lock.waitLock(30000);
  } catch (err) {
    logError_(new Error('lock 대기 시간 초과'), e);
    return json_({ ok: false, error: 'busy' });
  }

  try {
    const data = parse_(e);
    if (!data) throw new Error('본문을 해석할 수 없습니다');

    const sh = sheet_(SHEET_NAME, BASE_HEADERS);
    const headers = syncHeaders_(sh, Object.keys(data));

    // 같은 응답이 두 번 들어오면 두 번째는 버린다.
    // (페이지가 id 를 함께 보낼 때만 동작한다)
    if (isDuplicate_(sh, headers, data.id)) {
      return json_({ ok: true, duplicate: true });
    }

    sh.appendRow(headers.map(function (h) {
      return safe_(Object.prototype.hasOwnProperty.call(data, h) ? data[h] : '');
    }));

    return json_({ ok: true });

  } catch (err) {
    logError_(err, e);
    return json_({ ok: false, error: String((err && err.message) || err) });

  } finally {
    lock.releaseLock();
  }
}


/**
 * 웹앱 URL 이 살아 있는지 브라우저에서 눌러 확인하는 용도.
 * {"ok":true,"alive":true,...} 가 보이면 배포가 정상이다.
 */
function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME);
  return json_({
    ok: true,
    alive: true,
    sheet: SHEET_NAME,
    rows: sh ? Math.max(sh.getLastRow() - 1, 0) : 0
  });
}


/* ══════════════════════════════════════════════════════════════
   내부 도구
   ══════════════════════════════════════════════════════════════ */

/** 본문을 객체로 만든다. JSON 이 우선이고, 폼 전송 형태도 받아준다. */
function parse_(e) {
  if (e && e.postData && e.postData.contents) {
    try {
      return JSON.parse(e.postData.contents);
    } catch (err) {
      // JSON 이 아니면 아래 폼 파라미터를 시도한다
    }
  }
  if (e && e.parameter && Object.keys(e.parameter).length) {
    return e.parameter;
  }
  return null;
}


/**
 * 시트의 열이 모자라면 늘린다.
 *
 * 새로 만든 시트는 기본 26열이고, 시트는 필요하다고 열을 알아서 늘려주지
 * 않는다. 26열 밖을 getRange 하거나 열 수보다 긴 행을 appendRow 하면 그
 * 자리에서 예외가 난다. 이 설문은 OCI-R-K 를 포함하면 58열이라 반드시 넘는다.
 */
function ensureColumns_(sh, need) {
  const have = sh.getMaxColumns();
  if (need > have) sh.insertColumnsAfter(have, need - have);
}


/** 시트를 가져오고, 없으면 머리글을 세워서 만든다. */
function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers && headers.length) {
      ensureColumns_(sh, headers.length);
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
    sh.setFrozenRows(1);
  }
  return sh;
}


/**
 * 들어온 응답에 새 항목이 있으면 머리글 오른쪽에 열을 늘린다.
 * 나중에 OCI-R-K 18문항을 추가해도 이 스크립트를 고칠 필요가 없다.
 */
function syncHeaders_(sh, keys) {
  const lastCol = sh.getLastColumn();
  let headers = lastCol > 0
    ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String)
    : [];

  // 오른쪽 끝의 빈 칸은 열로 치지 않는다
  while (headers.length && headers[headers.length - 1] === '') headers.pop();

  const missing = keys.filter(function (k) { return headers.indexOf(k) === -1; });
  if (missing.length) {
    ensureColumns_(sh, headers.length + missing.length);
    sh.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
    headers = headers.concat(missing);
    sh.setFrozenRows(1);
  }
  return headers;
}


/** id 열에 같은 값이 이미 있는지 본다. */
function isDuplicate_(sh, headers, id) {
  if (!id) return false;
  const col = headers.indexOf('id') + 1;
  const last = sh.getLastRow();
  if (col === 0 || last < 2) return false;

  const ids = sh.getRange(2, col, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return true;
  }
  return false;
}


/**
 * =, +, -, @ 로 시작하는 글자는 시트가 수식으로 읽어버린다.
 * 자유 응답(17-1, 21번)에 그런 글자가 들어올 수 있으므로 텍스트로 고정한다.
 * 앞에 붙는 작은따옴표는 화면에 보이지 않는다.
 */
function safe_(v) {
  if (typeof v === 'string' && /^[=+\-@]/.test(v)) return "'" + v;
  return v;
}


/** 실패를 _오류 시트에 남긴다. 여기서 또 실패해도 제출 자체를 막지는 않는다. */
function logError_(err, e) {
  try {
    const sh = sheet_(ERROR_SHEET, ['시각', '오류', '받은 본문']);
    sh.appendRow([
      new Date(),
      String((err && err.stack) || err),
      (e && e.postData && e.postData.contents)
        ? String(e.postData.contents).slice(0, 5000)
        : ''
    ]);
  } catch (ignore) {
    // 오류 기록에 실패하면 더 할 수 있는 일이 없다
  }
}


/** JSON 으로 응답한다. */
function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


/* ══════════════════════════════════════════════════════════════
   점검용 — 편집기에서 직접 실행해본다
   ══════════════════════════════════════════════════════════════ */

/**
 * 설문 페이지 없이 가짜 응답 한 건을 넣어본다.
 * 편집기 상단에서 이 함수를 골라 실행하면 응답 시트에 한 줄이 생긴다.
 * 확인 후 그 줄은 지우면 된다.
 */
function testInsert() {
  const res = doPost({
    postData: {
      contents: JSON.stringify({
        id: 'test-' + Date.now(),
        ts: new Date().toISOString(),
        weeklyMin: 232,
        ref: 'test',
        f0: '하루 여러 번',
        f6: '거의 매일',
        q10: '2~3회',
        q14: '약속이나 수업에 늦거나 가지 못함 / 잠들기 어려웠거나 잠을 설침',
        q17b: '=참아보기',
        q26: '예상보다 훨씬 많다'
      })
    }
  });
  Logger.log(res.getContent());
}


/**
 * 응답 시트를 비운다. 머리글은 남긴다.
 * 시범 운영 뒤 실제 조사를 시작하기 전에 한 번 쓰는 용도다.
 * 되돌릴 수 없으니 실행 전에 사본을 떠둘 것.
 */
function resetResponses() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sh) return;
  const last = sh.getLastRow();
  if (last > 1) sh.deleteRows(2, last - 1);
}
